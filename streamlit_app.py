import streamlit as st
import os
import time
import json
from datetime import datetime
from pathlib import Path
import uuid
import base64
from typing import Dict, Any

# Import our presentation generator
from app.ppt_gen.ppt_generator_agent import generate_powerpoint_presentation

# Configure the Streamlit page
st.set_page_config(
    page_title="AI Presentation Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #2E86AB;
        margin-bottom: 2rem;
    }
    .feature-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .status-success {
        color: #28a745;
        font-weight: bold;
    }
    .status-error {
        color: #dc3545;
        font-weight: bold;
    }
    .status-warning {
        color: #ffc107;
        font-weight: bold;
    }
    .stProgress > div > div > div > div {
        background-color: #2E86AB;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Initialize session state variables"""
    if 'presentation_result' not in st.session_state:
        st.session_state.presentation_result = None
    if 'generation_in_progress' not in st.session_state:
        st.session_state.generation_in_progress = False
    if 'thread_id' not in st.session_state:
        st.session_state.thread_id = str(uuid.uuid4())[:8]

def get_binary_file_downloader_html(bin_file, file_label='File'):
    """Generate download link for binary files"""
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        bin_str = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">📥 Download {file_label}</a>'
        return href
    except Exception as e:
        return f"Error creating download link: {str(e)}"

def display_generation_progress():
    """Display progress indicators during presentation generation"""
    progress_steps = [
        "🔍 Researching topic and gathering content...",
        "📝 Planning presentation structure...",
        "🎨 Generating PowerPoint code...",
        "✅ Validating generated code...",
        "🚀 Creating presentation file...",
        "📧 Sending email (if requested)..."
    ]
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, step in enumerate(progress_steps):
        progress_bar.progress((i + 1) / len(progress_steps))
        status_text.text(step)
        time.sleep(1)  # Simulated delay for demonstration
    
    return progress_bar, status_text

def display_result_summary(result: Dict[Any, Any]):
    """Display a comprehensive summary of the generation result"""
    if result['success']:
        st.success("🎉 Presentation generated successfully!")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Generation Summary")
            st.write(f"**Topic:** {result['topic']}")
            st.write(f"**Thread ID:** {result['thread_id']}")
            
            if result.get('ppt_path'):
                st.write(f"**File Location:** {result['ppt_path']}")
                
                # Check if the ppt_path exists and provide download
                ppt_path = result.get('ppt_path')
                if ppt_path and Path(ppt_path).exists():
                    st.markdown(get_binary_file_downloader_html(
                        ppt_path, 
                        f"Presentation: {Path(ppt_path).name}"
                    ), unsafe_allow_html=True)
                else:
                    # Try to find the actual PPTX file in generated presentations
                    gen_folder = Path("generated_presentations")
                    if gen_folder.exists():
                        pptx_files = list(gen_folder.rglob("*.pptx"))
                        if pptx_files:
                            # Show the most recent PPTX file
                            latest_pptx = max(pptx_files, key=os.path.getmtime)
                            st.markdown(get_binary_file_downloader_html(
                                str(latest_pptx), 
                                f"Latest Presentation: {latest_pptx.name}"
                            ), unsafe_allow_html=True)
                            st.info(f"📍 Found presentation at: {latest_pptx}")
                        else:
                            st.warning("⚠️ No presentation files found in the generated folder")
            
            if result.get('email_sent'):
                st.write("📧 **Email Status:** ✅ Sent successfully")
            elif result.get('email_recipient'):
                st.write("📧 **Email Status:** ❌ Failed to send")
        
        with col2:
            st.subheader("🔄 Generation Process")
            if result.get('workflow_log'):
                for i, log_entry in enumerate(result['workflow_log'][-5:]):  # Show last 5 steps
                    agent_name = log_entry.get('agent', 'Unknown')
                    st.write(f"**Step {i+1}:** {agent_name}")
                    st.caption(log_entry.get('content', '')[:100] + "..." if len(log_entry.get('content', '')) > 100 else log_entry.get('content', ''))
    else:
        st.error("❌ Presentation generation failed")
        st.write(f"**Error:** {result.get('error', 'Unknown error occurred')}")
        
        # Show workflow log for debugging
        if result.get('workflow_log'):
            with st.expander("🔍 View generation log for debugging"):
                for log_entry in result['workflow_log']:
                    st.write(f"**{log_entry.get('agent', 'Unknown')}:**")
                    st.code(log_entry.get('content', ''), language='text')
        
        # Additional debugging info
        with st.expander("🛠️ Technical Details"):
            st.json(result)

def main():
    """Main Streamlit application"""
    initialize_session_state()
    
    # Header
    st.markdown("<h1 class='main-header'>🤖 AI-Powered Presentation Generator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>Create professional PowerPoint presentations using AI agents</p>", unsafe_allow_html=True)
    
    # Sidebar configuration
    with st.sidebar:
        st.header("🛠️ Configuration")
        
        # Environment check
        st.subheader("Environment Status")
        env_status = []
        
        # Check for required environment variables
        required_vars = ['OPENAI_API_KEY', 'SEARCHAPI_API_KEY']
        for var in required_vars:
            if os.getenv(var):
                env_status.append(f"✅ {var}")
            else:
                env_status.append(f"❌ {var}")
        
        for status in env_status:
            if "✅" in status:
                st.markdown(f"<span class='status-success'>{status}</span>", unsafe_allow_html=True)
            else:
                st.markdown(f"<span class='status-error'>{status}</span>", unsafe_allow_html=True)
        
        st.divider()
        
        # Advanced options
        st.subheader("Advanced Options")
        show_workflow_details = st.checkbox("Show detailed workflow", value=False)
        auto_download = st.checkbox("Auto-download generated presentations", value=True)
    
    # Main content area
    tab1, tab2, tab3 = st.tabs(["🎯 Generate Presentation", "📁 Recent Presentations", "ℹ️ About"])
    
    with tab1:
        # Presentation generation form
        st.subheader("📝 Create Your Presentation")
        
        with st.form("presentation_form"):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                topic = st.text_area(
                    "Enter your presentation topic:",
                    placeholder="e.g., 'The Future of Artificial Intelligence in Healthcare' or 'Digital Marketing Strategies for Small Businesses'",
                    height=100,
                    help="Be specific and descriptive for better results"
                )
            
            with col2:
                st.markdown("### 💡 Tips for Better Results")
                st.markdown("""
                - Be specific about your topic
                - Include target audience context
                - Mention key points you want covered
                - Specify the presentation style if needed
                """)
            
            # Email configuration
            st.subheader("📧 Email Delivery (Optional)")
            email_recipient = st.text_input(
                "Recipient email address:",
                placeholder="colleague@company.com",
                help="Leave empty to skip email delivery"
            )
            
            # Submit button
            submitted = st.form_submit_button(
                "🚀 Generate Presentation",
                type="primary",
                use_container_width=True
            )
        
        # Handle form submission
        if submitted and topic.strip():
            if not st.session_state.generation_in_progress:
                st.session_state.generation_in_progress = True
                
                with st.spinner("🔄 Generating your presentation..."):
                    # Show progress
                    progress_container = st.container()
                    with progress_container:
                        st.info("🤖 AI agents are working on your presentation...")
                        
                        if show_workflow_details:
                            progress_bar, status_text = display_generation_progress()
                    
                    # Generate the presentation
                    try:
                        result = generate_powerpoint_presentation(
                            topic=topic.strip(),
                            thread_id=st.session_state.thread_id,
                            email_recipient=email_recipient.strip() if email_recipient.strip() else None
                        )
                        
                        st.session_state.presentation_result = result
                        
                    except Exception as e:
                        st.error(f"An error occurred during generation: {str(e)}")
                        st.session_state.presentation_result = {
                            'success': False,
                            'error': str(e),
                            'topic': topic
                        }
                    
                    finally:
                        st.session_state.generation_in_progress = False
        
        elif submitted and not topic.strip():
            st.warning("⚠️ Please enter a topic for your presentation.")
        
        # Display results
        if st.session_state.presentation_result:
            st.divider()
            display_result_summary(st.session_state.presentation_result)
    
    with tab2:
        st.subheader("📁 Recent Presentations")
        
        # List generated presentations
        gen_folder = Path("generated_presentations")
        if gen_folder.exists():
            pptx_files = list(gen_folder.rglob("*.pptx"))
            
            if pptx_files:
                st.write(f"Found {len(pptx_files)} presentations:")
                
                # Sort by modification time (newest first)
                pptx_files.sort(key=os.path.getmtime, reverse=True)
                
                for pptx_file in pptx_files[:10]:  # Show last 10
                    col1, col2, col3 = st.columns([3, 2, 1])
                    
                    with col1:
                        st.write(f"📊 **{pptx_file.name}**")
                        st.caption(f"Location: {pptx_file.parent}")
                    
                    with col2:
                        mod_time = datetime.fromtimestamp(os.path.getmtime(pptx_file))
                        st.write(f"🕒 {mod_time.strftime('%Y-%m-%d %H:%M')}")
                    
                    with col3:
                        st.markdown(get_binary_file_downloader_html(
                            str(pptx_file), 
                            pptx_file.name
                        ), unsafe_allow_html=True)
            else:
                st.info("No presentations generated yet. Create your first presentation in the 'Generate Presentation' tab!")
        else:
            st.info("No presentations folder found. Generate your first presentation to get started!")
    
    with tab3:
        st.subheader("ℹ️ About This Application")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 🤖 How It Works
            
            This application uses a multi-agent AI system to create professional PowerPoint presentations:
            
            1. **Content Planner Agent** - Researches your topic and plans the content structure
            2. **Code Generator Agent** - Creates Python code to build the presentation
            3. **Validator Agent** - Ensures the code is correct and functional
            4. **Executor Agent** - Runs the code and generates the final PPTX file
            5. **Email Agent** - Optionally sends the presentation via email
            
            ### 🛠️ Technologies Used
            - **LangGraph** for agent orchestration
            - **OpenAI GPT** for AI capabilities
            - **Python-PPTX** for presentation generation
            - **Streamlit** for the web interface
            """)
        
        with col2:
            st.markdown("""
            ### 🚀 Features
            
            - **AI-Powered Content**: Automatically researches and creates engaging content
            - **Professional Design**: Uses proven presentation templates and layouts
            - **Multi-Agent System**: Specialized AI agents for different tasks
            - **Email Integration**: Direct delivery to recipients
            - **Download Support**: Easy access to generated files
            - **Real-time Progress**: Track generation status
            
            ### 📋 Requirements
            
            To use this application, you need:
            - OpenAI API key
            - SearchAPI key (for web research)
            - Internet connection for research and image retrieval
            """)
        
        st.divider()
        
        # Footer
        st.markdown("""
        <div style='text-align: center; color: #666; margin-top: 2rem;'>
            Made with ❤️ using Streamlit and AI • © 2024 AI Presentation Generator
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
