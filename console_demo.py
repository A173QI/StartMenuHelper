#!/usr/bin/env python3
"""
Start Menu Shortcut Creator - Console Demo Version
This is a simulated console version of the Windows Start Menu Shortcut Creator
"""
import os
import sys
from shortcut_creator_demo import ShortcutCreator

def clear_screen():
    """Clear the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Print the application header."""
    print("\n" + "=" * 80)
    print(" " * 25 + "START MENU SHORTCUT CREATOR")
    print(" " * 20 + "Console Demo (Windows Simulation Mode)")
    print("=" * 80 + "\n")

def print_menu():
    """Print the main menu."""
    print("\nMAIN MENU:")
    print("1. Create a new shortcut")
    print("2. About this application")
    print("3. Exit")
    
    choice = input("\nEnter your choice (1-3): ")
    return choice

def get_exe_path():
    """Get the path to the executable file."""
    print("\nEnter the path to the executable file (.exe, .bat, .cmd, .msi)")
    print("Example: C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
    print("Demo note: You can enter any valid Windows path in simulation mode")
    
    path = input("\nPath: ")
    
    # For the demo, accept any input as long as it has a valid extension
    if not (path.lower().endswith('.exe') or 
            path.lower().endswith('.bat') or 
            path.lower().endswith('.cmd') or 
            path.lower().endswith('.msi')):
        print("\nWarning: File doesn't have a recognized Windows executable extension.")
        confirm = input("Do you want to continue anyway? (y/n): ")
        if confirm.lower() != 'y':
            return None
    
    return path

def get_shortcut_details(creator, exe_path):
    """Get details for the shortcut creation."""
    # Get executable info
    exe_info = creator.get_exe_info(exe_path)
    
    print("\nEXECUTABLE INFORMATION:")
    print(f"Path: {exe_path}")
    print(f"Product: {exe_info['product_name']} (v{exe_info['version']})")
    print(f"Description: {exe_info['description']}")
    print(f"Publisher: {exe_info['company']}")
    
    # Get shortcut name
    suggested_name = exe_info['suggested_name']
    print(f"\nSuggested shortcut name: {suggested_name}")
    name = input("Enter shortcut name (or press Enter to use suggested): ")
    if not name:
        name = suggested_name
    
    # Get install location
    print("\nINSTALL LOCATION:")
    print("1. Current User Start Menu (no admin rights needed)")
    print("2. All Users Start Menu (requires admin privileges)")
    
    location_choice = input("Enter choice (1-2): ")
    for_all_users = (location_choice == '2')
    
    if for_all_users and not creator.is_admin():
        print("\nWARNING: Creating shortcuts for all users requires administrator privileges.")
        print("In a real Windows environment, you would need to restart the application as Administrator.")
        print("For this demo, we'll continue in simulation mode.")
    
    # Get subfolder
    subfolder = input("\nEnter optional subfolder (or press Enter for none): ")
    if not subfolder:
        subfolder = None
    
    return name, for_all_users, subfolder

def create_shortcut():
    """Handle the shortcut creation process."""
    clear_screen()
    print_header()
    print("\nCREATE NEW SHORTCUT")
    print("-" * 80)
    
    creator = ShortcutCreator()
    
    # Get executable path
    exe_path = get_exe_path()
    if not exe_path:
        print("\nShortcut creation cancelled.")
        input("\nPress Enter to return to main menu...")
        return
    
    # Get shortcut details
    name, for_all_users, subfolder = get_shortcut_details(creator, exe_path)
    
    # Preview and confirm
    clear_screen()
    print_header()
    print("\nSHORTCUT PREVIEW")
    print("-" * 80)
    
    print(f"Executable: {exe_path}")
    print(f"Shortcut Name: {name}.lnk")
    
    # Determine destination path
    if for_all_users:
        dest_base = "All Users Start Menu"
        dest_path = creator.common_start_menu
    else:
        dest_base = "Current User Start Menu"
        dest_path = creator.user_start_menu
    
    if subfolder:
        dest = f"{dest_base} → {subfolder}"
        dest_path = os.path.join(dest_path, subfolder)
    else:
        dest = dest_base
    
    print(f"Destination: {dest}")
    print(f"Full path: {os.path.join(dest_path, name + '.lnk')}")
    
    # Confirm creation
    confirm = input("\nCreate this shortcut? (y/n): ")
    if confirm.lower() != 'y':
        print("\nShortcut creation cancelled.")
        input("\nPress Enter to return to main menu...")
        return
    
    # Create the shortcut
    success, message = creator.create_shortcut(exe_path, name, for_all_users, subfolder)
    
    if success:
        print(f"\nSUCCESS: {message}")
    else:
        print(f"\nERROR: {message}")
    
    input("\nPress Enter to return to main menu...")

def show_about():
    """Show information about the application."""
    clear_screen()
    print_header()
    print("\nABOUT THIS APPLICATION")
    print("-" * 80)
    print("\nStart Menu Shortcut Creator")
    print("Version: 1.0.0 (Demo)")
    print("\nThis is a demonstration of the Windows Start Menu Shortcut Creator")
    print("running in a simulated environment.")
    print("\nIn a real Windows environment, this application would:")
    print(" - Select Windows executable (.exe) files")
    print(" - Create shortcuts in the Windows Start Menu")
    print(" - Support both Current User and All Users (admin) installations")
    print(" - Allow organizing shortcuts into subfolders")
    print("\nThis demo simulates these features but doesn't actually create shortcuts.")
    
    input("\nPress Enter to return to main menu...")

def demo_mode():
    """Run a demonstration of both functions automatically."""
    # Show the about information
    clear_screen()
    print_header()
    print("\nDEMO MODE - Automatic demonstration of all features")
    print("-" * 80)
    print("\nPart 1: About this application")
    print("-" * 40)
    
    print("\nStart Menu Shortcut Creator")
    print("Version: 1.0.0 (Demo)")
    print("\nThis is a demonstration of the Windows Start Menu Shortcut Creator")
    print("running in a simulated environment.")
    print("\nIn a real Windows environment, this application would:")
    print(" - Select Windows executable (.exe) files")
    print(" - Create shortcuts in the Windows Start Menu")
    print(" - Support both Current User and All Users (admin) installations")
    print(" - Allow organizing shortcuts into subfolders")
    print("\nThis demo simulates these features but doesn't actually create shortcuts.")
    
    print("\n" + "-" * 80)
    print("\nPart 2: Creating a shortcut")
    print("-" * 40)
    
    # Demo the shortcut creation process
    creator = ShortcutCreator()
    
    # Use a sample executable path
    exe_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    print(f"\nSelected executable: {exe_path}")
    
    # Get executable info
    exe_info = creator.get_exe_info(exe_path)
    
    print("\nEXECUTABLE INFORMATION:")
    print(f"Path: {exe_path}")
    print(f"Product: {exe_info['product_name']} (v{exe_info['version']})")
    print(f"Description: {exe_info['description']}")
    print(f"Publisher: {exe_info['company']}")
    
    # Use the suggested shortcut name
    name = exe_info['suggested_name']
    print(f"\nShortcut name: {name}")
    
    # Choose current user installation
    for_all_users = False
    print("\nInstall location: Current User Start Menu")
    
    # No subfolder
    subfolder = None
    
    # Show preview
    print("\nSHORTCUT PREVIEW")
    print("-" * 40)
    print(f"Executable: {exe_path}")
    print(f"Shortcut Name: {name}.lnk")
    
    # Determine destination path
    dest_base = "Current User Start Menu"
    dest_path = creator.user_start_menu
    dest = dest_base
    
    print(f"Destination: {dest}")
    print(f"Full path: {os.path.join(dest_path, name + '.lnk')}")
    
    # Create the shortcut
    print("\nCreating the shortcut...")
    success, message = creator.create_shortcut(exe_path, name, for_all_users, subfolder)
    
    if success:
        print(f"\nSUCCESS: {message}")
    else:
        print(f"\nERROR: {message}")
    
    print("\n" + "-" * 80)
    print("\nDemo completed. This shows how the application would work in a real Windows environment.")
    print("In the full version, users would have more options and could create actual shortcuts.")
    
def main():
    """Main application function."""
    # Run in demo mode directly
    demo_mode()
    
    print("\nThank you for trying the Start Menu Shortcut Creator Demo!")
    sys.exit(0)

if __name__ == "__main__":
    main()