"""
Image Integration Module for PowerPoint Presentations
Provides tools for searching, downloading, and managing images from Unsplash API
"""

import os
import requests
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from PIL import Image
import tempfile
import uuid
from datetime import datetime

# Configure logging
logger = logging.getLogger(__name__)

class UnsplashImageManager:
    """Manages image search and download from Unsplash API"""
    
    def __init__(self, access_key: Optional[str] = None):
        """
        Initialize Unsplash image manager
        
        Args:
            access_key: Unsplash API access key (if None, will try to get from environment)
        """
        self.access_key = access_key or os.getenv('UNSPLASH_ACCESS_KEY')
        if not self.access_key:
            logger.warning("No Unsplash access key provided. Image search will not work.")
        
        self.base_url = "https://api.unsplash.com"
        self.headers = {
            "Authorization": f"Client-ID {self.access_key}" if self.access_key else "",
            "Accept-Version": "v1"
        }
    
    def search_images(self, query: str, per_page: int = 10, orientation: str = "landscape") -> List[Dict]:
        """
        Search for images on Unsplash
        
        Args:
            query: Search query for images
            per_page: Number of results to return (max 30)
            orientation: Image orientation ('landscape', 'portrait', 'squarish')
            
        Returns:
            List of image data dictionaries
        """
        if not self.access_key:
            logger.error("No Unsplash access key available for image search")
            return []
        
        try:
            url = f"{self.base_url}/search/photos"
            params = {
                "query": query,
                "per_page": min(per_page, 30),
                "orientation": orientation,
                "order_by": "relevance"
            }
            
            response = requests.get(url, headers=self.headers, params=params, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            results = data.get('results', [])
            
            # Extract relevant information
            images = []
            for result in results:
                image_info = {
                    "id": result.get("id"),
                    "description": result.get("description") or result.get("alt_description", ""),
                    "url_regular": result.get("urls", {}).get("regular"),
                    "url_small": result.get("urls", {}).get("small"),
                    "url_thumb": result.get("urls", {}).get("thumb"),
                    "width": result.get("width"),
                    "height": result.get("height"),
                    "photographer": result.get("user", {}).get("name"),
                    "photographer_url": result.get("user", {}).get("links", {}).get("html"),
                    "download_url": result.get("links", {}).get("download_location")
                }
                images.append(image_info)
            
            logger.info(f"Found {len(images)} images for query: {query}")
            return images
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error searching images: {str(e)}")
            return []
        except Exception as e:
            logger.error(f"Unexpected error during image search: {str(e)}")
            return []
    
    def download_image(self, image_url: str, save_path: str, max_size: Tuple[int, int] = (1920, 1080)) -> bool:
        """
        Download and optionally resize an image
        
        Args:
            image_url: URL of the image to download
            save_path: Local path to save the image
            max_size: Maximum size (width, height) for resizing
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Download the image
            response = requests.get(image_url, timeout=30)
            response.raise_for_status()
            
            # Save to temporary file first
            with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp_file:
                temp_file.write(response.content)
                temp_path = temp_file.name
            
            # Open and process with PIL
            with Image.open(temp_path) as img:
                # Convert to RGB if necessary (for PNG with transparency)
                if img.mode in ('RGBA', 'LA', 'P'):
                    rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    rgb_img.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                    img = rgb_img
                
                # Resize if image is larger than max_size
                if img.size[0] > max_size[0] or img.size[1] > max_size[1]:
                    img.thumbnail(max_size, Image.Resampling.LANCZOS)
                
                # Ensure save directory exists
                Path(save_path).parent.mkdir(parents=True, exist_ok=True)
                
                # Save the processed image
                img.save(save_path, 'JPEG', quality=85, optimize=True)
            
            # Clean up temp file
            os.unlink(temp_path)
            
            logger.info(f"Image downloaded and saved to: {save_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error downloading image: {str(e)}")
            # Clean up temp file if it exists
            try:
                if 'temp_path' in locals():
                    os.unlink(temp_path)
            except:
                pass
            return False
    
    def get_best_image_for_topic(self, topic: str, project_folder: str) -> Optional[Dict]:
        """
        Find and download the best image for a given topic
        
        Args:
            topic: Topic to search for
            project_folder: Project folder to save images
            
        Returns:
            Dictionary with image information if successful, None otherwise
        """
        try:
            # Search for images
            images = self.search_images(topic, per_page=5)
            
            if not images:
                logger.warning(f"No images found for topic: {topic}")
                return None
            
            # Select the best image (first result, which is most relevant)
            best_image = images[0]
            
            # Generate filename
            image_id = best_image.get("id", str(uuid.uuid4())[:8])
            filename = f"image_{image_id}.jpg"
            
            # Set up save path
            assets_folder = Path(project_folder) / "assets"
            save_path = assets_folder / filename
            
            # Download the image
            image_url = best_image.get("url_regular") or best_image.get("url_small")
            if image_url and self.download_image(image_url, str(save_path)):
                best_image["local_path"] = str(save_path)
                best_image["filename"] = filename
                return best_image
            else:
                logger.error(f"Failed to download image for topic: {topic}")
                return None
                
        except Exception as e:
            logger.error(f"Error getting best image for topic '{topic}': {str(e)}")
            return None

class FallbackImageManager:
    """Fallback image manager for when Unsplash API is not available"""
    
    def __init__(self):
        self.placeholder_colors = [
            "#3498db",  # Blue
            "#2ecc71",  # Green
            "#e74c3c",  # Red
            "#f39c12",  # Orange
            "#9b59b6",  # Purple
            "#1abc9c",  # Turquoise
        ]
    
    def create_placeholder_image(self, text: str, size: Tuple[int, int] = (800, 600), 
                                save_path: str = None) -> str:
        """
        Create a placeholder image with text
        
        Args:
            text: Text to display on the image
            size: Image size (width, height)
            save_path: Path to save the image
            
        Returns:
            Path to the created image
        """
        try:
            from PIL import Image, ImageDraw, ImageFont
            
            # Create image
            color = self.placeholder_colors[hash(text) % len(self.placeholder_colors)]
            img = Image.new('RGB', size, color)
            draw = ImageDraw.Draw(img)
            
            # Try to use a nice font, fall back to default
            try:
                font = ImageFont.truetype("arial.ttf", 48)
            except:
                font = ImageFont.load_default()
            
            # Calculate text position
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            x = (size[0] - text_width) // 2
            y = (size[1] - text_height) // 2
            
            # Draw text
            draw.text((x, y), text, fill="white", font=font)
            
            # Save image
            if not save_path:
                save_path = f"placeholder_{uuid.uuid4().hex[:8]}.jpg"
            
            Path(save_path).parent.mkdir(parents=True, exist_ok=True)
            img.save(save_path, 'JPEG', quality=85)
            
            logger.info(f"Created placeholder image: {save_path}")
            return save_path
            
        except Exception as e:
            logger.error(f"Error creating placeholder image: {str(e)}")
            return ""

# Factory function for image manager
def create_image_manager() -> UnsplashImageManager:
    """Create an image manager instance"""
    return UnsplashImageManager()

# Utility functions for integration with existing tools
def search_and_download_image(topic: str, project_folder: str, 
                            use_fallback: bool = True) -> Optional[Dict]:
    """
    Search and download an image for a given topic
    
    Args:
        topic: Topic to search for
        project_folder: Project folder to save images
        use_fallback: Whether to create placeholder if search fails
        
    Returns:
        Image information dictionary or None
    """
    manager = create_image_manager()
    result = manager.get_best_image_for_topic(topic, project_folder)
    
    if not result and use_fallback:
        # Create fallback placeholder
        fallback_manager = FallbackImageManager()
        assets_folder = Path(project_folder) / "assets"
        placeholder_path = assets_folder / f"placeholder_{uuid.uuid4().hex[:8]}.jpg"
        
        created_path = fallback_manager.create_placeholder_image(
            topic.title(), 
            save_path=str(placeholder_path)
        )
        
        if created_path:
            result = {
                "id": "placeholder",
                "description": f"Placeholder for {topic}",
                "local_path": created_path,
                "filename": Path(created_path).name,
                "is_placeholder": True
            }
    
    return result

if __name__ == "__main__":
    # Test the image manager
    print("Testing Unsplash Image Manager...")
    
    # Create test project folder
    test_folder = Path("test_images")
    test_folder.mkdir(exist_ok=True)
    
    # Test image search and download
    manager = create_image_manager()
    images = manager.search_images("artificial intelligence", per_page=3)
    
    if images:
        print(f"Found {len(images)} images")
        best_image = manager.get_best_image_for_topic("artificial intelligence", str(test_folder))
        if best_image:
            print(f"Downloaded image: {best_image['local_path']}")
        else:
            print("Failed to download image")
    else:
        print("No images found or API key not configured")
        
        # Test fallback
        fallback = FallbackImageManager()
        placeholder_path = test_folder / "test_placeholder.jpg"
        created = fallback.create_placeholder_image("AI Technology", save_path=str(placeholder_path))
        print(f"Created placeholder: {created}")




