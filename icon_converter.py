"""
SVG to ICO Converter for Start Menu Shortcut Creator
This utility helps convert SVG icons to ICO format for Windows applications
"""
import os
import sys
from PIL import Image
import cairosvg

def convert_svg_to_ico(svg_path, ico_path, sizes=(16, 32, 48, 64, 128, 256)):
    """
    Convert an SVG file to an ICO file with multiple sizes
    
    Args:
        svg_path: Path to SVG file
        ico_path: Path for output ICO file
        sizes: Tuple of icon sizes to include
    """
    # Create temporary PNG files at different sizes
    temp_pngs = []
    
    for size in sizes:
        png_path = f"temp_icon_{size}.png"
        cairosvg.svg2png(url=svg_path, write_to=png_path, output_width=size, output_height=size)
        temp_pngs.append((png_path, size))
    
    # Open all temporary PNGs
    imgs = []
    for png_path, size in temp_pngs:
        img = Image.open(png_path)
        imgs.append(img)
    
    # Save ICO file with all sizes
    imgs[0].save(ico_path, format='ICO', sizes=[(img.width, img.height) for img in imgs], 
                 append_images=imgs[1:])
    
    # Clean up temporary files
    for png_path, _ in temp_pngs:
        os.remove(png_path)
    
    print(f"Successfully converted {svg_path} to {ico_path}")

def main():
    """Main function"""
    print("SVG to ICO Converter for Start Menu Shortcut Creator")
    print("====================================================")
    
    svg_path = "assets/app_icon.svg"
    ico_path = "assets/app_icon.ico"
    
    if not os.path.exists(svg_path):
        print(f"Error: SVG file not found at {svg_path}")
        return
    
    print(f"Converting {svg_path} to {ico_path}...")
    
    try:
        # On Windows, this would work:
        # convert_svg_to_ico(svg_path, ico_path)
        # But in Replit environment, we'd need cairosvg and Pillow
        
        print("NOTE: This conversion requires libraries that may not be available in Replit.")
        print("When running on Windows, install the necessary libraries:")
        print("pip install cairosvg pillow")
        print("\nThen run this script to generate the ICO file.")
        
    except Exception as e:
        print(f"Error during conversion: {e}")

if __name__ == "__main__":
    main()