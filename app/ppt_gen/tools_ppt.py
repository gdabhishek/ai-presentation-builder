
from langchain_core.tools import tool

from langchain_community.tools.arxiv.tool import ArxivQueryRun
from langgraph_supervisor import create_supervisor
from langchain_community.utilities import SearchApiAPIWrapper
from langchain.tools import Tool
from langchain_community.tools.wikipedia.tool import WikipediaQueryRun
from langchain_community.utilities import WikipediaAPIWrapper
from langchain_community.utilities import ArxivAPIWrapper
# Initialize research tools

import logging
logger = logging.getLogger(__name__)
import logging
import os
import uuid
import subprocess
import sys
import ast
import tempfile
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()
# search_tool = DuckDuckGoSearchRun()



wikipedia = WikipediaQueryRun(api_wrapper=WikipediaAPIWrapper())
arxiv_wrapper = ArxivAPIWrapper(top_k_results=2, doc_content_chars_max=1000)
search = SearchApiAPIWrapper()

search_tool = Tool(
    name="web_search",
    func=search.run,
    description="Search the web using SearchApi"
)
arxiv_tool = ArxivQueryRun(arxiv_wrapper=arxiv_wrapper)

@tool
def create_project_folder(topic: str) -> dict:
    """
    Create a unique project folder for this presentation generation request.
    
    Args:
        topic (str): The topic for the presentation
        
    Returns:
        dict: Contains unique_id and project_folder path
    """
    try:
        unique_id = str(uuid.uuid4())[:8]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"ppt_presentation_{topic}_{timestamp}_{unique_id}"
        project_folder = Path.cwd() / "generated_presentations" / folder_name
        project_folder.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Created project folder: {project_folder}")
        
        # Create subdirectories
        (project_folder / "code").mkdir(exist_ok=True)
        (project_folder / "presentations").mkdir(exist_ok=True)
        (project_folder / "assets").mkdir(exist_ok=True)
        (project_folder / "logs").mkdir(exist_ok=True)
        
        return {
            "unique_id": unique_id,
            "project_folder": str(project_folder),
            "status": "success"
        }
    except Exception as e:
        logger.error(f"Error creating project folder: {str(e)}")
        return {
            "unique_id": "",
            "project_folder": "",
            "status": "error",
            "error": str(e)
        }

@tool
def check_pptx_installation() -> dict:
    """
    Check if python-pptx is installed and install if necessary.
    
    Returns:
        dict: Installation status and version info
    """
    try:
        # Check if python-pptx is installed
        result = subprocess.run([sys.executable, "-c", "import pptx; print(pptx.__version__)"], 
                              capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            version = result.stdout.strip()
            logger.info(f"python-pptx is already installed, version: {version}")
            return {
                "status": "already_installed",
                "version": version,
                "message": f"python-pptx version {version} is ready to use"
            }
        else:
            # Install python-pptx
            logger.info("Installing python-pptx...")
            install_result = subprocess.run([sys.executable, "-m", "pip", "install", "python-pptx"], 
                                          capture_output=True, text=True, timeout=120)
            
            if install_result.returncode == 0:
                # Verify installation
                verify_result = subprocess.run([sys.executable, "-c", "import pptx; print(pptx.__version__)"], 
                                             capture_output=True, text=True, timeout=30)
                
                if verify_result.returncode == 0:
                    version = verify_result.stdout.strip()
                    logger.info(f"python-pptx successfully installed, version: {version}")
                    return {
                        "status": "newly_installed",
                        "version": version,
                        "message": f"python-pptx version {version} has been installed successfully"
                    }
                else:
                    logger.error("python-pptx installation verification failed")
                    return {
                        "status": "error",
                        "message": "python-pptx installation verification failed"
                    }
            else:
                error_msg = install_result.stderr or "Unknown installation error"
                logger.error(f"python-pptx installation failed: {error_msg}")
                return {
                    "status": "error",
                    "message": f"python-pptx installation failed: {error_msg}"
                }
                
    except subprocess.TimeoutExpired:
        logger.error("python-pptx installation/check timed out")
        return {
            "status": "error",
            "message": "python-pptx installation/check timed out"
        }
    except Exception as e:
        logger.error(f"Error checking/installing python-pptx: {str(e)}")
        return {
            "status": "error",
            "message": f"Error checking/installing python-pptx: {str(e)}"
        }

@tool
def validate_ppt_code(code: str) -> dict:
    """
    Validate Python code for PowerPoint generation.
    
    Args:
        code (str): Python code to validate
        
    Returns:
        dict: Validation results
    """
    try:
        logger.info(f"Validating code: {code}")
        # Parse the code to check for syntax errors
        ast.parse(code)
        
        validation_results = {
            "syntax_valid": True,
            "imports_template": False,
            "uses_template_system": False,
            "creates_presentation": False,
            "saves_presentation": False,
            "has_slides": False,
            "template_theme_selected": False,
            "uses_raw_pptx": False,
            "errors": [],
            "warnings": []
        }
        
        # Check for PowerPoint-specific requirements
        lines = code.split('\n')
        code_text = '\n'.join(lines)
        
        # Check for template import (REQUIRED)
        if 'from ppt_templates import create_template' in code_text:
            validation_results["imports_template"] = True
        
        # Check for template usage (REQUIRED)
        if 'create_template(' in code_text:
            validation_results["uses_template_system"] = True
            validation_results["creates_presentation"] = True
            
        # Check for template theme selection
        themes = ["business", "tech", "creative", "corporate"]
        for theme in themes:
            if f'create_template("{theme}")' in code_text:
                validation_results["template_theme_selected"] = True
                break
        
        # Check for raw pptx usage (NOT ALLOWED)
        if ('Presentation()' in code_text or 'pptx.Presentation()' in code_text or
            'from pptx import Presentation' in code_text or 'import pptx' in code_text):
            validation_results["uses_raw_pptx"] = True
        
        # Check for saving presentation
        if '.save(' in code_text:
            validation_results["saves_presentation"] = True
        
        # Check for slide creation using template methods
        template_methods = ["create_title_slide", "create_content_slide", "create_section_slide", 
                          "create_comparison_slide", "create_conclusion_slide", "create_thank_you_slide",
                          "create_image_slide", "create_image_content_slide", "create_image_comparison_slide"]
        for method in template_methods:
            if method in code_text:
                validation_results["has_slides"] = True
                break
        
        # Check for image path references in bullet points (common mistake)
        problematic_patterns = [
            "Visual note:",
            "Image suggestion:",
            "supporting image",
            "image located at",
            "assets/image_",
            "project_images/",
            ".jpg",
            ".png"
        ]
        
        for pattern in problematic_patterns:
            if pattern in code_text and ("create_content_slide" in code_text):
                validation_results["warnings"].append(
                    f"WARNING: Found '{pattern}' in code - images should use image slide methods, not text references"
                )
        
        # Add errors/warnings for template requirements
        if not validation_results["imports_template"]:
            validation_results["errors"].append("CRITICAL: Missing template import. Must include: from ppt_templates import create_template")
        
        if not validation_results["uses_template_system"]:
            validation_results["errors"].append("CRITICAL: Must use template system. Include: ppt = create_template('theme_name')")
        
        if not validation_results["template_theme_selected"]:
            validation_results["warnings"].append("Template theme not clearly specified. Use: business, tech, creative, or corporate")
        
        if validation_results["uses_raw_pptx"]:
            validation_results["errors"].append("CRITICAL: Raw python-pptx usage detected. Must use template system only")
        
        if not validation_results["saves_presentation"]:
            validation_results["warnings"].append("Presentation save method not found")
        
        if not validation_results["has_slides"]:
            validation_results["warnings"].append("No template slide creation methods found")
        
        return validation_results
        
    except SyntaxError as e:
        return {
            "syntax_valid": False,
            "imports_pptx": False,
            "creates_presentation": False,
            "saves_presentation": False,
            "has_slides": False,
            "errors": [f"Syntax error at line {e.lineno}: {e.msg}"],
            "warnings": []
        }
    except Exception as e:
        return {
            "syntax_valid": False,
            "imports_pptx": False,
            "creates_presentation": False,
            "saves_presentation": False,
            "has_slides": False,
            "errors": [f"Validation error: {str(e)}"],
            "warnings": []
        }


@tool
def save_code_to_file(code: str, project_folder: str, filename: str = "presentation.pptx") -> dict:
    '''
    Save the code to a file.
    
    Args:
        code (str): Python code to save
        project_folder (str): Project folder path
        filename (str): Name of the file to save
        
    Returns:
        dict: Execution results
    '''
    try:
        project_path = Path(project_folder)
        code_file = project_path / "code" / "generate_ppt.py"
        with open(code_file, 'w', encoding='utf-8') as f:
            f.write(code)

        return {
            "status": "success",
            "message": f"Code saved to {code_file}"
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Error saving code to file: {str(e)}"
        }

@tool
def execute_ppt_code(code: str, project_folder: str, filename: str = "presentation.pptx") -> dict:
    """
    Execute PowerPoint generation code and create the presentation.
    
    Args:
        code (str): Python code to execute
        project_folder (str): Project folder path
        filename (str): Name of the PowerPoint file
        
    Returns:
        dict: Execution results
    """
    try:
        project_path = Path(project_folder)
        code_file = project_path / "code" / "generate_ppt.py"
        ppt_output_dir = project_path / "presentations"
        
        # Ensure filename has .pptx extension
        if not filename.endswith('.pptx'):
            filename += '.pptx'
        
        # Modify code to save in the correct location
        output_path = ppt_output_dir / filename
        
        # Fix save path in the code
        output_path_str = str(output_path).replace('\\', '/')  # Use forward slashes for cross-platform compatibility
        
        if '.save(' in code:
            # Replace existing save calls with correct path
            import re
            # Find and replace save calls, handle both quoted and unquoted arguments
            code = re.sub(r'\.save\([^)]*\)', f'.save(r"{output_path_str}")', code)
        else:
            # Add save command if missing
            code += f'\n\n# Save the presentation\nprs.save(r"{output_path_str}")'
        
        # Always copy template file to project directory for template-based code
        if 'from ppt_templates import' in code:
            template_source = r"app\ppt_gen\ppt_templates.py"
            template_dest = project_path / "ppt_templates.py"
            if os.path.exists(template_source):
                import shutil
                shutil.copy2(template_source, template_dest)
                logger.info(f"Copied ppt_templates.py to {template_dest}")
            else:
                logger.error(f"Template source file not found at {template_source}")
                return {
                    "status": "error",
                    "ppt_path": "",
                    "message": f"Template source file not found at {template_source}"
                }
        
        # Also copy template to code directory for easier import
        if 'from ppt_templates import' in code:
            code_template_dest = project_path / "code" / "ppt_templates.py"
            if template_source.exists():
                import shutil
                shutil.copy2(template_source, code_template_dest)
                logger.info(f"Copied ppt_templates.py to {code_template_dest}")
        
        # Add Python path setup to code if using templates
        if 'from ppt_templates import' in code:
            code = f"""import sys
import os
# Add current directory to Python path for template imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

{code}"""
        
        # Write code to file
        with open(code_file, 'w', encoding='utf-8') as f:
            f.write(code)
        
        logger.info(f"Executing PowerPoint code in {code_file}")
        
        # Execute the code
        result = subprocess.run([sys.executable, str(code_file)], 
                              capture_output=True, text=True, timeout=120, cwd=str(project_path))
        
        if result.returncode == 0:
            # Check if presentation file was created
            if output_path.exists():
                ppt_path = str(output_path)
                logger.info(f"Presentation generated successfully: {ppt_path}")
                return {
                    "status": "success",
                    "ppt_path": ppt_path,
                    "output": result.stdout,
                    "message": "PowerPoint presentation generated successfully"
                }
            else:
                logger.warning("Code execution completed but no presentation file found")
                return {
                    "status": "warning",
                    "ppt_path": "",
                    "output": result.stdout,
                    "message": "Code execution completed but no presentation file found"
                }
        else:
            error_msg = result.stderr or result.stdout
            logger.error(f"PowerPoint code execution failed: {error_msg}")
            return {
                "status": "error",
                "ppt_path": "",
                "output": result.stdout,
                "error": error_msg,
                "message": f"PowerPoint code execution failed: {error_msg}"
            }
            
    except subprocess.TimeoutExpired:
        logger.error("PowerPoint code execution timed out")
        return {
            "status": "error",
            "ppt_path": "",
            "message": "PowerPoint code execution timed out (>2 minutes)"
        }
    except Exception as e:
        logger.error(f"Error executing PowerPoint code: {str(e)}")
        return {
            "status": "error",
            "ppt_path": "",
            "message": f"Error executing PowerPoint code: {str(e)}"
        }

