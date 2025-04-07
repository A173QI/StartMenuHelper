"""
Start Menu Shortcut Creator - Installer Creator
This script creates a simple installer using NSIS
"""
import os
import sys
import shutil
from pathlib import Path

def create_nsis_script(exe_path, version="1.0.0", output_dir="installer"):
    """
    Create an NSIS script for building a Windows installer
    
    Args:
        exe_path: Path to the .exe file to include
        version: Version number as string
        output_dir: Directory to save the script
    
    Returns:
        Path to the created NSIS script
    """
    # Get application name from exe filename
    app_name = os.path.basename(exe_path).replace(".exe", "")
    app_name_underscores = app_name.replace(" ", "_")
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Path to the NSIS script
    script_path = os.path.join(output_dir, f"{app_name_underscores}_installer.nsi")
    
    # NSIS script content
    script_content = f"""
; Installer script for {app_name}
Unicode True

; Basic definitions
!define APPNAME "{app_name}"
!define APPVERSION "{version}"
!define PUBLISHER "Replit"
!define WEBSITE "https://replit.com/"

; Include modern UI
!include "MUI2.nsh"

; General settings
Name "${{APPNAME}} ${{APPVERSION}}"
OutFile "${{APPNAME}}_${{APPVERSION}}_Setup.exe"
InstallDir "$PROGRAMFILES\\${{APPNAME}}"
InstallDirRegKey HKLM "Software\\${{APPNAME}}" "Install_Dir"
RequestExecutionLevel admin

; Interface settings
!define MUI_ABORTWARNING
!define MUI_ICON "assets\\app_icon.ico"
!define MUI_UNICON "assets\\app_icon.ico"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer sections
Section "Install"
    SetOutPath "$INSTDIR"
    
    ; Add files
    File "{exe_path}"
    File "assets\\app_icon.ico"
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\\uninstall.exe"
    
    ; Create shortcuts
    CreateDirectory "$SMPROGRAMS\\${{APPNAME}}"
    CreateShortcut "$SMPROGRAMS\\${{APPNAME}}\\${{APPNAME}}.lnk" "$INSTDIR\\{os.path.basename(exe_path)}" "" "$INSTDIR\\app_icon.ico"
    CreateShortcut "$SMPROGRAMS\\${{APPNAME}}\\Uninstall.lnk" "$INSTDIR\\uninstall.exe"
    CreateShortcut "$DESKTOP\\${{APPNAME}}.lnk" "$INSTDIR\\{os.path.basename(exe_path)}" "" "$INSTDIR\\app_icon.ico"
    
    ; Add registry keys for uninstall
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayName" "${{APPNAME}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "UninstallString" "$INSTDIR\\uninstall.exe"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayIcon" "$INSTDIR\\app_icon.ico"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "Publisher" "${{PUBLISHER}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "DisplayVersion" "${{APPVERSION}}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "URLInfoAbout" "${{WEBSITE}}"
    
    ; Write info about install size
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
    IntFmt $0 "0x%08X" $0
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}" "EstimatedSize" "$0"
SectionEnd

; Uninstaller section
Section "Uninstall"
    ; Remove files and uninstaller
    Delete "$INSTDIR\\{os.path.basename(exe_path)}"
    Delete "$INSTDIR\\app_icon.ico"
    Delete "$INSTDIR\\uninstall.exe"
    
    ; Remove shortcuts
    Delete "$SMPROGRAMS\\${{APPNAME}}\\${{APPNAME}}.lnk"
    Delete "$SMPROGRAMS\\${{APPNAME}}\\Uninstall.lnk"
    Delete "$DESKTOP\\${{APPNAME}}.lnk"
    RMDir "$SMPROGRAMS\\${{APPNAME}}"
    
    ; Remove registry keys
    DeleteRegKey HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${{APPNAME}}"
    DeleteRegKey HKLM "Software\\${{APPNAME}}"
    
    ; Remove install directory
    RMDir "$INSTDIR"
SectionEnd
"""

    # Create LICENSE.txt file if it doesn't exist
    license_path = os.path.join(output_dir, "LICENSE.txt")
    if not os.path.exists(license_path):
        with open(license_path, "w") as f:
            f.write(f"""MIT License

Copyright (c) {os.environ.get('YEAR', '2023')} {os.environ.get('PUBLISHER', 'Replit')}

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
""")
    
    # Write the NSIS script to file
    with open(script_path, "w") as f:
        f.write(script_content)
    
    # Copy icon file to output directory
    icon_source = "assets/app_icon.ico"
    icon_dest = os.path.join(output_dir, "app_icon.ico")
    
    if os.path.exists(icon_source):
        shutil.copy(icon_source, icon_dest)
    else:
        print(f"Warning: Icon file not found at {icon_source}")
        
    return script_path

def main():
    print("Start Menu Shortcut Creator - Installer Creator")
    print("==============================================")
    
    # Look for executable in dist directory
    dist_dir = "dist"
    exe_files = []
    
    if os.path.exists(dist_dir):
        exe_files = [f for f in os.listdir(dist_dir) if f.endswith('.exe')]
    
    if not exe_files:
        print("No .exe files found in the 'dist' directory.")
        print("Please build the executable first using PyInstaller or build_windows_exe.bat")
        return 1
    
    # Select the executable to use
    if len(exe_files) == 1:
        exe_path = os.path.join(dist_dir, exe_files[0])
    else:
        print("Multiple .exe files found. Please select one:")
        for i, exe in enumerate(exe_files):
            print(f"{i+1}. {exe}")
        
        selection = input("Enter the number of the executable to use: ")
        try:
            index = int(selection) - 1
            if 0 <= index < len(exe_files):
                exe_path = os.path.join(dist_dir, exe_files[index])
            else:
                print("Invalid selection.")
                return 1
        except ValueError:
            print("Invalid input.")
            return 1
    
    # Create NSIS script
    output_dir = "installer"
    script_path = create_nsis_script(exe_path, output_dir=output_dir)
    
    print(f"\nNSIS script created at: {script_path}")
    print("\nTo build the installer:")
    print("1. Install NSIS from https://nsis.sourceforge.io/")
    print("2. Right-click on the .nsi file and select 'Compile NSIS Script'")
    print("   OR run: makensis.exe installer/*.nsi")
    print(f"\nThe installer will be created in the '{output_dir}' directory.")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())