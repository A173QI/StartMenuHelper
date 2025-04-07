# Start Menu Shortcut Creator

A Windows desktop application that simplifies creating Start Menu shortcuts with a user-friendly interface.

## Features

- Drag and drop interface for selecting executable files
- Automatic extraction of application information and icons
- Dynamic icon preview with selectable sizes (16×16, 32×32, 48×48)
- Support for custom shortcut names
- Options for current user or all users installation
- Organization of shortcuts into subfolders
- Administrator privilege elevation when needed
- Shortcut verification and repair tools

## How to Use

1. Start the application - Run the executable file after installation
2. Select an executable - Drag and drop a .exe file or click to browse
3. Configure shortcut - Modify the shortcut name if desired
4. Choose installation location - Current user or All users (requires admin)
5. Optional: Add subfolder - Organize shortcuts by specifying a subfolder
6. Create shortcut - Click the "Create Shortcut" button

## Requirements

- Windows 7 or newer
- Administrator rights (for "All Users" installation)

## Project Structure

- main.py - Main application entry point with PyQt GUI
- shortcut_creator.py - Core Windows shortcut functionality
- shortcut_creator_demo.py - Simulated shortcut creation for testing
- console_demo.py - Console-based demo of the application
- ui_components.py - UI components for the PyQt application
- styles.py - Style definitions for the UI
- icon_extractor.py - Extract application icons from executables
- shortcut_verifier.py - Verify and repair broken shortcuts
- icon_converter.py - Tool to convert SVG icons to ICO format
- setup.py - Setup script for distribution
- create_installer.py - Tool to create a Windows installer
- build_windows_exe.bat - Batch file to build on Windows
- assets/ - Application icons and images
