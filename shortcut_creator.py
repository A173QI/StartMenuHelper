import os
import sys
import ctypes
import winreg
import pythoncom
from win32com.client import Dispatch
import win32com.shell.shell as shell
import win32com.shell.shellcon as shellcon
import win32api
import win32con
import traceback

class ShortcutCreator:
    def __init__(self):
        self.common_start_menu = self._get_common_start_menu_path()
        self.user_start_menu = self._get_user_start_menu_path()
    
    def _get_user_start_menu_path(self):
        """Get the path to the current user's Start Menu Programs folder."""
        try:
            return shell.SHGetFolderPath(0, shellcon.CSIDL_PROGRAMS, 0, 0)
        except Exception:
            # Default fallback path
            return os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
    
    def _get_common_start_menu_path(self):
        """Get the path to the All Users Start Menu Programs folder."""
        try:
            return shell.SHGetFolderPath(0, shellcon.CSIDL_COMMON_PROGRAMS, 0, 0)
        except Exception:
            # Default fallback path
            return os.path.join(os.environ["PROGRAMDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
    
    def is_admin(self):
        """Check if the current process has admin privileges."""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False

    def is_valid_exe(self, file_path):
        """Verify that the file is a valid Windows executable."""
        if not os.path.exists(file_path):
            return False
        
        # Check file extension
        if not file_path.lower().endswith(('.exe', '.bat', '.cmd', '.msi')):
            return False
        
        # For .exe files, do additional verification
        if file_path.lower().endswith('.exe'):
            try:
                # Check if it has a valid PE header
                with open(file_path, 'rb') as f:
                    header = f.read(2)
                    return header == b'MZ'  # Valid PE files start with 'MZ'
            except Exception:
                return False
        
        return True
    
    def get_exe_info(self, exe_path):
        """Extract version info from an executable."""
        try:
            # Get file info
            info = win32api.GetFileVersionInfo(exe_path, '\\')
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            
            # Get string file info
            string_file_info = {}
            try:
                lang, codepage = win32api.GetFileVersionInfo(exe_path, '\\VarFileInfo\\Translation')[0]
                str_info_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\%s'
                
                for info_name in ['FileDescription', 'ProductName', 'CompanyName']:
                    try:
                        string_file_info[info_name] = win32api.GetFileVersionInfo(exe_path, str_info_path % info_name)
                    except:
                        string_file_info[info_name] = ""
            except:
                pass
            
            return {
                'version': version,
                'description': string_file_info.get('FileDescription', ''),
                'product_name': string_file_info.get('ProductName', ''),
                'company': string_file_info.get('CompanyName', ''),
                'suggested_name': string_file_info.get('ProductName', os.path.splitext(os.path.basename(exe_path))[0])
            }
        except Exception as e:
            # If we can't get version info, just return basic info
            basename = os.path.splitext(os.path.basename(exe_path))[0]
            return {
                'version': '',
                'description': '',
                'product_name': '',
                'company': '',
                'suggested_name': basename
            }
    
    def create_shortcut(self, exe_path, shortcut_name, for_all_users=False, folder=None):
        """
        Create a shortcut to the executable in the Start Menu.
        
        Args:
            exe_path: Path to the executable
            shortcut_name: Name for the shortcut
            for_all_users: If True, create in All Users Start Menu (requires admin)
            folder: Optional subfolder within Start Menu Programs
            
        Returns:
            (success, message) tuple
        """
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Select the appropriate Start Menu path
            if for_all_users:
                if not self.is_admin():
                    return False, "Administrator privileges required to create shortcuts for all users."
                start_menu_path = self.common_start_menu
            else:
                start_menu_path = self.user_start_menu
            
            # Create target folder if specified
            if folder:
                target_dir = os.path.join(start_menu_path, folder)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)
            else:
                target_dir = start_menu_path
            
            # Ensure valid shortcut name
            if not shortcut_name.endswith('.lnk'):
                shortcut_name += '.lnk'
            
            # Full path to the shortcut
            shortcut_path = os.path.join(target_dir, shortcut_name)
            
            # Create the shortcut
            shell_object = Dispatch('WScript.Shell')
            shortcut = shell_object.CreateShortCut(shortcut_path)
            shortcut.Targetpath = exe_path
            shortcut.WorkingDirectory = os.path.dirname(exe_path)
            shortcut.IconLocation = f"{exe_path},0"  # Use first icon from the exe
            shortcut.save()
            
            return True, f"Shortcut created successfully at:\n{shortcut_path}"
            
        except Exception as e:
            error_msg = str(e)
            detailed_error = traceback.format_exc()
            return False, f"Error creating shortcut: {error_msg}\n\nDetails:\n{detailed_error}"
        
        finally:
            # Clean up COM
            pythoncom.CoUninitialize()
