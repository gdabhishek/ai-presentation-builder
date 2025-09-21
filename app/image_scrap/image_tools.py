"""
LangChain tools for image integration in PowerPoint presentations
"""

import logging
from typing import Dict, Optional
from langchain_core.tools import tool
from pathlib import Path
from app.image_scrap.image_integration import search_and_download_image, create_image_manager

logger = logging.getLogger(__name__)

@tool
def search_presentation_image(topic: str, project_folder: str) -> dict:
    """
    Search and download an appropriate image for a presentation topic.
    
    Args:
        topic (str): The topic or subject to search images for
        project_folder (str): Project folder path where images should be saved
        
    Returns:
        dict: Image information including local path, description, and metadata
    """
    try:
        # Search and download image
        image_info = search_and_download_image(topic, project_folder, use_fallback=True)
        
        if image_info:
            return {
                "status": "success",
                "image_id": image_info.get("id", "unknown"),
                "local_path": image_info.get("local_path", ""),
                "filename": image_info.get("filename", ""),
                "description": image_info.get("description", f"Image for {topic}"),
                "photographer": image_info.get("photographer", ""),
                "is_placeholder": image_info.get("is_placeholder", False),
                "message": f"Successfully found and downloaded image for '{topic}'"
            }
        else:
            return {
                "status": "error",
                "image_id": "",
                "local_path": "",
                "filename": "",
                "description": "",
                "photographer": "",
                "is_placeholder": False,
                "message": f"Failed to find or download image for '{topic}'"
            }
            
    except Exception as e:
        logger.error(f"Error in search_presentation_image: {str(e)}")
        return {
            "status": "error",
            "image_id": "",
            "local_path": "",
            "filename": "",
            "description": "",
            "photographer": "",
            "is_placeholder": False,
            "message": f"Error searching for image: {str(e)}"
        }

@tool
def search_multiple_images(topics: str, project_folder: str, max_images: int = 5) -> dict:
    """
    Search and download multiple images for different topics in a presentation.
    
    Args:
        topics (str): Comma-separated list of topics to search images for
        project_folder (str): Project folder path where images should be saved
        max_images (int): Maximum number of images to download
        
    Returns:
        dict: Results for all image searches including paths and metadata
    """
    try:
        topic_list = [topic.strip() for topic in topics.split(',')]
        topic_list = topic_list[:max_images]  # Limit to max_images
        
        results = []
        successful_downloads = 0
        
        for topic in topic_list:
            if not topic:
                continue
                
            image_info = search_and_download_image(topic, project_folder, use_fallback=True)
            
            if image_info:
                results.append({
                    "topic": topic,
                    "status": "success",
                    "image_id": image_info.get("id", "unknown"),
                    "local_path": image_info.get("local_path", ""),
                    "filename": image_info.get("filename", ""),
                    "description": image_info.get("description", f"Image for {topic}"),
                    "photographer": image_info.get("photographer", ""),
                    "is_placeholder": image_info.get("is_placeholder", False)
                })
                successful_downloads += 1
            else:
                results.append({
                    "topic": topic,
                    "status": "error",
                    "image_id": "",
                    "local_path": "",
                    "filename": "",
                    "description": "",
                    "photographer": "",
                    "is_placeholder": False
                })
        
        return {
            "status": "success" if successful_downloads > 0 else "error",
            "total_requested": len(topic_list),
            "successful_downloads": successful_downloads,
            "results": results,
            "message": f"Downloaded {successful_downloads} out of {len(topic_list)} requested images"
        }
        
    except Exception as e:
        logger.error(f"Error in search_multiple_images: {str(e)}")
        return {
            "status": "error",
            "total_requested": 0,
            "successful_downloads": 0,
            "results": [],
            "message": f"Error searching for multiple images: {str(e)}"
        }

@tool 
def get_image_suggestions(topic: str) -> dict:
    """
    Get image search suggestions and previews for a presentation topic without downloading.
    
    Args:
        topic (str): The topic to get image suggestions for
        
    Returns:
        dict: List of suggested images with metadata
    """
    try:
        manager = create_image_manager()
        images = manager.search_images(topic, per_page=8)
        
        if images:
            suggestions = []
            for img in images:
                suggestions.append({
                    "id": img.get("id", ""),
                    "description": img.get("description", ""),
                    "photographer": img.get("photographer", ""),
                    "url_thumb": img.get("url_thumb", ""),
                    "url_small": img.get("url_small", ""),
                    "dimensions": f"{img.get('width', 0)}x{img.get('height', 0)}"
                })
            
            return {
                "status": "success",
                "topic": topic,
                "total_found": len(suggestions),
                "suggestions": suggestions,
                "message": f"Found {len(suggestions)} image suggestions for '{topic}'"
            }
        else:
            return {
                "status": "warning",
                "topic": topic,
                "total_found": 0,
                "suggestions": [],
                "message": f"No image suggestions found for '{topic}'"
            }
            
    except Exception as e:
        logger.error(f"Error in get_image_suggestions: {str(e)}")
        return {
            "status": "error",
            "topic": topic,
            "total_found": 0,
            "suggestions": [],
            "message": f"Error getting image suggestions: {str(e)}"
        }

@tool
def validate_image_setup() -> dict:
    """
    Validate that image integration is properly configured.
    
    Returns:
        dict: Validation results including API key status and dependencies
    """
    try:
        import os
        import requests
        from PIL import Image
        
        validation_results = {
            "dependencies_installed": True,
            "unsplash_api_key": bool(os.getenv('UNSPLASH_ACCESS_KEY')),
            "can_process_images": True,
            "can_make_requests": True,
            "errors": [],
            "warnings": []
        }
        
        # Check Unsplash API key
        if not validation_results["unsplash_api_key"]:
            validation_results["warnings"].append(
                "Unsplash API key not found. Set UNSPLASH_ACCESS_KEY environment variable for image search."
            )
        
        # Test basic functionality
        try:
            manager = create_image_manager()
            if validation_results["unsplash_api_key"]:
                # Test API connection with a simple search
                test_images = manager.search_images("test", per_page=1)
                if not test_images and manager.access_key:
                    validation_results["warnings"].append(
                        "Unsplash API connection test failed. Check API key validity."
                    )
        except Exception as e:
            validation_results["errors"].append(f"Image manager test failed: {str(e)}")
        
        # Determine overall status
        if validation_results["errors"]:
            status = "error"
            message = "Image integration has errors that need to be fixed"
        elif validation_results["warnings"]:
            status = "warning"
            message = "Image integration is functional but has warnings"
        else:
            status = "success"
            message = "Image integration is fully configured and ready"
        
        return {
            "status": status,
            "message": message,
            **validation_results
        }
        
    except ImportError as e:
        return {
            "status": "error",
            "message": f"Missing required dependencies: {str(e)}",
            "dependencies_installed": False,
            "unsplash_api_key": False,
            "can_process_images": False,
            "can_make_requests": False,
            "errors": [f"Import error: {str(e)}"],
            "warnings": []
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Validation error: {str(e)}",
            "dependencies_installed": False,
            "unsplash_api_key": False,
            "can_process_images": False,
            "can_make_requests": False,
            "errors": [str(e)],
            "warnings": []
        }

