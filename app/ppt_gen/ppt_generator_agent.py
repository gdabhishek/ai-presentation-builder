# Environment setup
from dotenv import load_dotenv
load_dotenv()

import logging
import os
import uuid
import subprocess
import sys
import ast
import tempfile
from pathlib import Path
from typing import Annotated, Sequence, TypedDict, Optional
from datetime import datetime

# LangGraph and LangChain imports
from langchain_openai import ChatOpenAI
from langchain_anthropic import ChatAnthropic
from langgraph.graph import StateGraph, END, START
from langchain_core.messages import BaseMessage, HumanMessage
from langgraph.graph.message import add_messages

from langgraph.prebuilt import create_react_agent
from langgraph.checkpoint.memory import InMemorySaver
from langgraph_supervisor import create_supervisor

from app.ppt_gen.tools_ppt import search_tool, wikipedia, arxiv_tool, create_project_folder, check_pptx_installation, execute_ppt_code, validate_ppt_code, save_code_to_file
from app.image_scrap.image_tools import search_presentation_image, search_multiple_images, get_image_suggestions, validate_image_setup
from app.ppt_gen.prompts import content_planner_prompt, ppt_code_generator_prompt, code_validator_prompt, ppt_executor_prompt,supervisor_prompt
from app.email.email_prompts import email_agent_prompt    
from app.email.email_tools import email_tools




logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ppt_generator.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class PPTGeneratorState(TypedDict):
    """Agent state for the PowerPoint generation workflow"""
    messages: Annotated[Sequence[BaseMessage], add_messages]
    topic: str
    research_content: str
    content_plan: str
    slide_structure: str
    ppt_code: str
    validation_result: str
    execution_result: str
    ppt_path: str
    unique_id: str
    project_folder: str
    errors: list
    current_iteration: int
    max_iterations: int


# Initialize LLMs
# llm_planner = ChatAnthropic(model="claude-sonnet-4-20250514")
# llm_generator = ChatAnthropic(model="claude-sonnet-4-20250514")
# llm_validator = ChatAnthropic(model="claude-sonnet-4-20250514")
# llm_executor = ChatAnthropic(model="claude-sonnet-4-20250514")
llm_supervisor = ChatAnthropic(model="claude-sonnet-4-20250514")
llm_planner = ChatOpenAI(model="gpt-5-mini")
llm_generator = ChatAnthropic(model="claude-sonnet-4-20250514")
llm_validator = ChatOpenAI(model="gpt-5-mini")
llm_executor = ChatOpenAI(model="gpt-5-mini")
llm_email = ChatOpenAI(model="gpt-5-mini")



image_tools = [
    search_presentation_image,
    search_multiple_images, 
    get_image_suggestions,
    validate_image_setup
]
content_planner_tools = [search_tool, wikipedia, arxiv_tool, create_project_folder] + image_tools
ppt_generator_tools = []  # No tools needed, just generates code
validator_tools = [validate_ppt_code]
executor_tools = [check_pptx_installation, execute_ppt_code]
email_sender_tools = email_tools  # Email sending tools

# Create agents
content_planner_agent = create_react_agent(
    llm_planner,
    content_planner_tools,
    prompt=content_planner_prompt,
    name="content_planner_agent"
)

ppt_code_generator_agent = create_react_agent(
    llm_generator,
    ppt_generator_tools,
    prompt=ppt_code_generator_prompt,
    name="ppt_code_generator_agent"
)

code_validator_agent = create_react_agent(
    llm_validator,
    validator_tools,
    prompt=code_validator_prompt,
    name="code_validator_agent"
)

ppt_executor_agent = create_react_agent(
    llm_executor,
    executor_tools,
    prompt=ppt_executor_prompt,
    name="ppt_executor_agent"
)

email_sender_agent = create_react_agent(
    llm_email,
    email_sender_tools,
    prompt=email_agent_prompt,
    name="email_sender_agent"
)

# Create the supervisor workflow
memory = InMemorySaver()
supervisor = create_supervisor(
    agents=[content_planner_agent, ppt_code_generator_agent, code_validator_agent, ppt_executor_agent, email_sender_agent],
    model=llm_supervisor,  
    prompt=(
        supervisor_prompt
    ),
    output_mode="full_history"
)
supervisor = supervisor.compile(checkpointer=memory)

def generate_powerpoint_presentation(topic: str, thread_id: str = None, email_recipient: str = None) -> dict:
    """
    Generate a PowerPoint presentation for the given topic using the supervisor workflow.
    
    Args:
        topic (str): The topic for the presentation
        thread_id (str): Thread ID for conversation tracking
        email_recipient (str): Optional email address to send the presentation to
        
    Returns:
        dict: Results including presentation path and execution details
    """
    if thread_id is None:
        thread_id = str(uuid.uuid4())[:8]
    
    logger.info(f"Starting PowerPoint presentation generation for topic: {topic}")
    logger.info("=" * 60)
    
    # Prepare input
    email_instruction = f" Also send the presentation to {email_recipient} via email." if email_recipient else ""
    input_messages = {
        "messages": [HumanMessage(content=f"Generate a professional PowerPoint presentation about: {topic}{email_instruction}")]
    }
    config = {"configurable": {"thread_id": thread_id}}
    
    # Track the workflow
    workflow_log = []
    agent_outputs = {}
    
    try:
        # Stream the supervisor execution
        for event in supervisor.stream(input_messages, config=config):
            for agent_name, value in event.items():
                if 'messages' in value and value['messages']:
                    latest_message = value['messages'][-1]
                    # Clean content for safe logging (replace problematic Unicode chars)
                    safe_content = latest_message.content.encode('ascii', 'replace').decode('ascii')
                    logger.info(f"Agent '{agent_name}': {safe_content}...")
                    logger.info("-" * 40)
                    
                    # Store for analysis
                    workflow_log.append({
                        "agent": agent_name,
                        "content": latest_message.content,
                        "timestamp": datetime.now().isoformat()
                    })
                    agent_outputs[agent_name] = latest_message.content
        
        # Analyze results
        result = {
            "success": False,
            "topic": topic,
            "thread_id": thread_id,
            "workflow_log": workflow_log,
            "agent_outputs": agent_outputs,
            "ppt_path": None,
            "project_folder": None,
            "email_sent": False,
            "email_recipient": email_recipient,
            "error": None
        }
        
        # Check if presentation was generated successfully
        executor_output = agent_outputs.get("ppt_executor_agent", "")
        logger.info(f"executor_output : {executor_output}")
        
        # Look for success indicators in the output
        success_indicators = [
            "PowerPoint presentation generated successfully",
            "ppt_path",
            "status.*success",
            "presentation.*created"
        ]
        
        is_successful = any(indicator.lower() in executor_output.lower() for indicator in success_indicators)
        
        if is_successful:
            result["success"] = True
            
            
        
        # Check if email was sent successfully
        email_output = agent_outputs.get("email_sender_agent", "")
        if email_output and ("Email sent successfully" in email_output or "✅" in email_output):
            result["email_sent"] = True
        
        logger.info(f"Presentation generation completed. Success: {result['success']}")
        return result
        
    except Exception as e:
        logger.error(f"Error during presentation generation: {str(e)}")
        return {
            "success": False,
            "topic": topic,
            "thread_id": thread_id,
            "error": str(e),
            "workflow_log": workflow_log
        }

if __name__ == "__main__":
    
    # Example usage
    logger.info("PowerPoint Presentation Generator Agent Swarm Ready")
    logger.info("=" * 60)
    
    # Interactive mode
    while True:
        try:
            topic = input("\nEnter a topic for presentation generation (or 'quit' to exit): ").strip()
            if topic.lower() in ['quit', 'exit', 'q']:
                break
            
            if not topic:
                print("Please enter a valid topic.")
                continue
            
            # Ask for optional email recipient
            email_recipient = input("Enter email address to send presentation to (optional, press Enter to skip): ").strip()
            if not email_recipient:
                email_recipient = None
            
            print(f"\nGenerating PowerPoint presentation for: {topic}")
            if email_recipient:
                print(f"Will send to: {email_recipient}")
            print("This may take a few minutes...")
            
            result = generate_powerpoint_presentation(topic, email_recipient=email_recipient)
            
            if result["success"]:
                print(f"Presentation generated successfully!")
                if result.get("ppt_path"):
                    print(f"Presentation saved at: {result['ppt_path']}")
                if result.get("email_sent"):
                    print(f"✅ Email sent successfully to: {result['email_recipient']}")
                elif result.get("email_recipient"):
                    print(f"❌ Failed to send email to: {result['email_recipient']}")
            else:
                print(f"Presentation generation failed: {result.get('error', 'Unknown error')}")
                
        except KeyboardInterrupt:
            print("\n\nExiting...")
            break
        except Exception as e:
            print(f"\nUnexpected error: {e}")
