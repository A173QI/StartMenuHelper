"""
Start Menu Shortcut Creator - Icon Extractor
This module provides functionality to extract icons from Windows executables
"""
import os
import sys
import tempfile
from pathlib import Path
from PIL import Image, ImageQt

# Demo mode constants for testing without Windows dependencies
DEMO_ICONS = {
    "chrome": "assets/demo_icons/chrome.svg",
    "word": "assets/demo_icons/word.svg",
    "excel": "assets/demo_icons/excel.svg",
    "vlc": "assets/demo_icons/vlc.svg",
    "unknown": "assets/demo_icons/unknown.svg",
    "default": "assets/app_icon.svg"
}

class IconExtractor:
    def __init__(self):
        """Initialize the IconExtractor with necessary settings."""
        self.temp_directory = tempfile.gettempdir()
        self.cache_directory = os.path.join(self.temp_directory, "icon_cache")
        
        # Create cache directory if it doesn't exist
        os.makedirs(self.cache_directory, exist_ok=True)
        
        # In demo mode, ensure demo icons directory exists
        if not self._is_windows():
            demo_dir = os.path.join("assets", "demo_icons")
            os.makedirs(demo_dir, exist_ok=True)

    def _is_windows(self):
        """Check if the system is Windows."""
        return sys.platform == "win32"

    def extract_icon(self, exe_path, size=32):
        """
        Extract an icon from an executable file.
        
        Args:
            exe_path: Path to the executable
            size: Desired icon size (default: 32)
            
        Returns:
            Path to the extracted icon file
        """
        if not exe_path or not os.path.exists(exe_path):
            return self._get_default_icon(size)
            
        # Generate a unique cache filename based on exe_path and size
        cache_name = f"{os.path.basename(exe_path)}_{size}.png"
        cache_path = os.path.join(self.cache_directory, cache_name)
        
        # If we already have this icon in cache, return it
        if os.path.exists(cache_path):
            return cache_path
            
        if self._is_windows():
            try:
                # On Windows, use win32 API to extract icon
                import win32ui
                import win32gui
                import win32con
                import win32api
                
                # Get the large icon
                large_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
                large_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
                
                # First try LoadImage
                try:
                    hicon = win32gui.LoadImage(0, exe_path, win32con.IMAGE_ICON, size, size, win32con.LR_LOADFROMFILE)
                except:
                    # If LoadImage fails, try ExtractIconEx
                    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
                    hicon = win32gui.ExtractIconEx(exe_path, 0, 1)[0][0]
                
                # Create a device context
                hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                hbmp = win32ui.CreateBitmap()
                hbmp.CreateCompatibleBitmap(hdc, size, size)
                hdc = hdc.CreateCompatibleDC()
                
                # Draw the icon onto the bitmap
                hdc.SelectObject(hbmp)
                hdc.DrawIcon((0, 0), hicon)
                
                # Convert bitmap to Python Image
                bmpinfo = hbmp.GetInfo()
                bmpstr = hbmp.GetBitmapBits(True)
                img = Image.frombuffer(
                    'RGBA', 
                    (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                    bmpstr, 'raw', 'BGRA', 0, 1
                )
                
                # Clean up resources
                win32gui.DestroyIcon(hicon)
                hdc.DeleteDC()
                
                # Save the image to the cache directory
                img.save(cache_path)
                return cache_path
                
            except Exception as e:
                print(f"Error extracting icon from {exe_path}: {e}")
                return self._get_default_icon(size)
        else:
            # In demo mode, return a simulated icon based on the filename
            return self._get_demo_icon(exe_path, size)

    def _get_demo_icon(self, exe_path, size=32):
        """Get a demo icon for non-Windows environments."""
        exe_name = os.path.basename(exe_path).lower()
        
        # Choose an appropriate demo icon based on the executable name
        if "chrome" in exe_name or "browser" in exe_name:
            demo_icon = DEMO_ICONS.get("chrome")
        elif "word" in exe_name:
            demo_icon = DEMO_ICONS.get("word")
        elif "excel" in exe_name or "spreadsheet" in exe_name:
            demo_icon = DEMO_ICONS.get("excel")
        elif "vlc" in exe_name or "player" in exe_name:
            demo_icon = DEMO_ICONS.get("vlc")
        else:
            demo_icon = DEMO_ICONS.get("unknown")
            
        # If the demo icon doesn't exist, use the default app icon
        if not demo_icon or not os.path.exists(demo_icon):
            return self._get_default_icon(size)
            
        return demo_icon

    def _get_default_icon(self, size=32):
        """Get the default icon when extraction fails."""
        default_icon = DEMO_ICONS.get("default")
        
        if not default_icon or not os.path.exists(default_icon):
            # If even the default icon is missing, create a blank one
            default_path = os.path.join(self.cache_directory, f"default_{size}.png")
            if not os.path.exists(default_path):
                img = Image.new('RGBA', (size, size), (200, 200, 200, 255))
                img.save(default_path)
            return default_path
            
        return default_icon

    def create_qt_icon(self, exe_path, size=32):
        """
        Create a PyQt icon from an executable file.
        
        Args:
            exe_path: Path to the executable
            size: Desired icon size
            
        Returns:
            QIcon object
        """
        from PyQt5.QtGui import QIcon, QPixmap
        
        icon_path = self.extract_icon(exe_path, size)
        return QIcon(QPixmap(icon_path))

    def get_all_icons(self, exe_path, sizes=(16, 32, 48)):
        """
        Extract multiple icon sizes from an executable.
        
        Args:
            exe_path: Path to the executable
            sizes: Tuple of desired sizes
            
        Returns:
            Dictionary of size -> icon path
        """
        result = {}
        for size in sizes:
            result[size] = self.extract_icon(exe_path, size)
        return result

def create_demo_icons():
    """Create demo icons if they don't exist."""
    # SVG demo icons are created manually and checked into the repository
    demo_dir = os.path.join("assets", "demo_icons")
    os.makedirs(demo_dir, exist_ok=True)

def main():
    """Main function for standalone testing."""
    print("Start Menu Shortcut Creator - Icon Extractor")
    print("===========================================")
    
    # Create demo icons
    create_demo_icons()
    
    extractor = IconExtractor()
    
    # Test with some common paths
    test_paths = [
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files\\Microsoft Office\\Office16\\WINWORD.EXE",
        "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe",
        "C:\\Windows\\System32\\notepad.exe"
    ]
    
    for path in test_paths:
        print(f"\nExtracting icon from {path}")
        icon_path = extractor.extract_icon(path)
        print(f"Icon extracted to: {icon_path}")
        
    print("\nExtracting multiple sizes from Chrome")
    sizes = extractor.get_all_icons(test_paths[0])
    for size, path in sizes.items():
        print(f"Size {size}px: {path}")

if __name__ == "__main__":
    main()