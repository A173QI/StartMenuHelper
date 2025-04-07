"""
Start Menu Shortcut Creator - Shortcut Verifier and Repair Tool
This module provides functionality to verify and repair Windows shortcuts
"""
import os
import sys
import time
from pathlib import Path

class ShortcutVerifier:
    def __init__(self):
        """Initialize the ShortcutVerifier with necessary paths and settings."""
        self.user_start_menu = self._get_user_start_menu_path()
        self.common_start_menu = self._get_common_start_menu_path()
        self.broken_shortcuts = []
        self.repaired_shortcuts = []
        self.verified_shortcuts = []

    def _get_user_start_menu_path(self):
        """Get the path to the current user's Start Menu Programs folder."""
        # In real Windows, we'd use Win32 API to get this path
        # For demo/testing, we'll use a simulated path
        if sys.platform == "win32":
            # On Windows, use actual Start Menu path
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            return shell.SpecialFolders("Programs")
        else:
            # On non-Windows, simulate the path
            return os.path.expanduser("~/.start_menu")

    def _get_common_start_menu_path(self):
        """Get the path to the All Users Start Menu Programs folder."""
        # In real Windows, we'd use Win32 API to get this path
        # For demo/testing, we'll use a simulated path
        if sys.platform == "win32":
            # On Windows, use actual Common Programs path
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            return shell.SpecialFolders("AllUsersPrograms")
        else:
            # On non-Windows, simulate the path
            return "/usr/local/share/applications"

    def is_admin(self):
        """Check if the current process has admin privileges."""
        if sys.platform == "win32":
            try:
                import ctypes
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            except:
                return False
        else:
            # On Unix-like systems, check if user is root
            return os.geteuid() == 0

    def find_shortcuts(self, location="both", subfolder=None):
        """
        Find all shortcuts in the Start Menu.
        
        Args:
            location: "user", "common", or "both"
            subfolder: Optional subfolder within the Start Menu
            
        Returns:
            List of shortcut paths
        """
        shortcuts = []
        
        locations = []
        if location in ["user", "both"]:
            locations.append(self.user_start_menu)
        if location in ["common", "both"]:
            if self.is_admin() or not sys.platform == "win32":
                locations.append(self.common_start_menu)
            else:
                print("Warning: Admin privileges required to access All Users shortcuts")
        
        for base_path in locations:
            search_path = base_path
            if subfolder:
                search_path = os.path.join(base_path, subfolder)
            
            if not os.path.exists(search_path):
                continue
                
            # In demo mode, we might not have actual shortcuts
            if sys.platform == "win32":
                # On real Windows, find .lnk files
                for root, dirs, files in os.walk(search_path):
                    for file in files:
                        if file.lower().endswith(".lnk"):
                            shortcuts.append(os.path.join(root, file))
            else:
                # In demo mode, create some simulated shortcuts
                self._create_demo_shortcuts(search_path)
                for root, dirs, files in os.walk(search_path):
                    for file in files:
                        if file.lower().endswith(".lnk"):
                            shortcuts.append(os.path.join(root, file))
                
        return shortcuts

    def _create_demo_shortcuts(self, path):
        """Create simulated shortcuts for demo purposes."""
        if not os.path.exists(path):
            os.makedirs(path, exist_ok=True)
            
        # Create some sample shortcuts for the demo
        apps = [
            {"name": "Google Chrome", "target": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "broken": False},
            {"name": "Notepad", "target": "C:\\Windows\\System32\\notepad.exe", "broken": False},
            {"name": "Broken App", "target": "C:\\Program Files\\NonExistent\\app.exe", "broken": True},
            {"name": "Microsoft Word", "target": "C:\\Program Files\\Microsoft Office\\Office16\\WINWORD.EXE", "broken": False},
            {"name": "Missing Game", "target": "D:\\Games\\MissingGame\\game.exe", "broken": True}
        ]
        
        for app in apps:
            shortcut_path = os.path.join(path, f"{app['name']}.lnk")
            if not os.path.exists(shortcut_path):
                with open(shortcut_path, "w") as f:
                    f.write(f"TARGET={app['target']}\nBROKEN={app['broken']}\n")

    def get_shortcut_target(self, shortcut_path):
        """
        Get the target path from a shortcut.
        
        Args:
            shortcut_path: Path to the shortcut file
            
        Returns:
            Target path or None if shortcut is invalid
        """
        if not os.path.exists(shortcut_path):
            return None
            
        if sys.platform == "win32":
            # On Windows, use actual shortcut reading
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                return shortcut.Targetpath
            except Exception as e:
                print(f"Error reading shortcut: {e}")
                return None
        else:
            # In demo mode, read from our simulated shortcut file
            try:
                with open(shortcut_path, "r") as f:
                    for line in f:
                        if line.startswith("TARGET="):
                            return line.strip().split("=", 1)[1]
                return None
            except Exception as e:
                print(f"Error reading demo shortcut: {e}")
                return None

    def is_target_valid(self, target_path):
        """
        Check if the target path exists and is valid.
        
        Args:
            target_path: Path to check
            
        Returns:
            True if target exists, False otherwise
        """
        if not target_path:
            return False
            
        if sys.platform == "win32":
            # On Windows, actually check if the file exists
            return os.path.exists(target_path)
        else:
            # In demo mode, check for "BROKEN=True" in our simulated data
            return not "NonExistent" in target_path and not "Missing" in target_path

    def verify_shortcut(self, shortcut_path):
        """
        Verify if a shortcut is valid (target exists).
        
        Args:
            shortcut_path: Path to the shortcut
            
        Returns:
            (is_valid, target_path, error_message)
        """
        target_path = self.get_shortcut_target(shortcut_path)
        
        if not target_path:
            return (False, None, "Unable to read shortcut target")
            
        is_valid = self.is_target_valid(target_path)
        
        if is_valid:
            self.verified_shortcuts.append((shortcut_path, target_path))
            return (True, target_path, None)
        else:
            self.broken_shortcuts.append((shortcut_path, target_path))
            return (False, target_path, "Target file does not exist")

    def verify_all_shortcuts(self, location="both", subfolder=None):
        """
        Verify all shortcuts in the Start Menu.
        
        Args:
            location: "user", "common", or "both"
            subfolder: Optional subfolder within the Start Menu
            
        Returns:
            (valid_count, broken_count, shortcuts_info)
        """
        shortcuts = self.find_shortcuts(location, subfolder)
        valid_count = 0
        broken_count = 0
        shortcuts_info = []
        
        for shortcut_path in shortcuts:
            is_valid, target_path, error_message = self.verify_shortcut(shortcut_path)
            shortcut_name = os.path.basename(shortcut_path)
            
            shortcuts_info.append({
                "name": shortcut_name,
                "path": shortcut_path,
                "target": target_path,
                "valid": is_valid,
                "error": error_message
            })
            
            if is_valid:
                valid_count += 1
            else:
                broken_count += 1
                
        return (valid_count, broken_count, shortcuts_info)

    def repair_shortcut(self, shortcut_path, new_target=None):
        """
        Attempt to repair a broken shortcut.
        
        Args:
            shortcut_path: Path to the shortcut to repair
            new_target: Optional new target path
            
        Returns:
            (success, message)
        """
        if not os.path.exists(shortcut_path):
            return (False, "Shortcut file not found")
            
        target_path = self.get_shortcut_target(shortcut_path) if new_target is None else new_target
        
        if not target_path:
            return (False, "Unable to determine target path")
            
        # In a real application, we would try to locate the moved file
        # For demo purposes, we'll simulate "finding" the correct path
        if sys.platform == "win32":
            try:
                # On Windows, actually update the shortcut
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                
                if new_target:
                    shortcut.Targetpath = new_target
                
                # If we have a new target, or the current target exists, save and return success
                if new_target or os.path.exists(shortcut.Targetpath):
                    shortcut.save()
                    self.repaired_shortcuts.append((shortcut_path, shortcut.Targetpath))
                    return (True, f"Shortcut repaired, now points to {shortcut.Targetpath}")
                
                # Try to find a similar path that exists
                if "Program Files" in shortcut.Targetpath:
                    # Try Program Files (x86) if original was in Program Files
                    alt_path = shortcut.Targetpath.replace("Program Files", "Program Files (x86)")
                    if os.path.exists(alt_path):
                        shortcut.Targetpath = alt_path
                        shortcut.save()
                        self.repaired_shortcuts.append((shortcut_path, alt_path))
                        return (True, f"Shortcut repaired, now points to {alt_path}")
                
                return (False, "Unable to locate the target application")
                
            except Exception as e:
                return (False, f"Error repairing shortcut: {e}")
        else:
            # In demo mode, simulate repairing the shortcut
            try:
                with open(shortcut_path, "r") as f:
                    content = f.readlines()
                
                if new_target:
                    # Update with new target
                    with open(shortcut_path, "w") as f:
                        for line in content:
                            if line.startswith("TARGET="):
                                f.write(f"TARGET={new_target}\n")
                            elif line.startswith("BROKEN="):
                                f.write("BROKEN=False\n")
                            else:
                                f.write(line)
                    self.repaired_shortcuts.append((shortcut_path, new_target))
                    return (True, f"Shortcut repaired, now points to {new_target}")
                else:
                    # Try to find a better target
                    old_target = None
                    for line in content:
                        if line.startswith("TARGET="):
                            old_target = line.strip().split("=", 1)[1]
                            break
                    
                    if old_target:
                        # Simulate finding a better path
                        if "NonExistent" in old_target:
                            new_target = old_target.replace("NonExistent", "Existent")
                        elif "MissingGame" in old_target:
                            new_target = "C:\\Program Files\\Steam\\steamapps\\common\\Game\\game.exe"
                        else:
                            return (False, "Unable to find a replacement target")
                            
                        # Update the shortcut
                        with open(shortcut_path, "w") as f:
                            for line in content:
                                if line.startswith("TARGET="):
                                    f.write(f"TARGET={new_target}\n")
                                elif line.startswith("BROKEN="):
                                    f.write("BROKEN=False\n")
                                else:
                                    f.write(line)
                        
                        self.repaired_shortcuts.append((shortcut_path, new_target))
                        return (True, f"Shortcut repaired, now points to {new_target}")
                
                return (False, "Unable to repair the shortcut")
                
            except Exception as e:
                return (False, f"Error repairing demo shortcut: {e}")

    def repair_all_shortcuts(self):
        """
        Attempt to repair all broken shortcuts.
        
        Returns:
            (success_count, failed_count, results)
        """
        success_count = 0
        failed_count = 0
        results = []
        
        for shortcut_path, target_path in self.broken_shortcuts:
            success, message = self.repair_shortcut(shortcut_path)
            shortcut_name = os.path.basename(shortcut_path)
            
            results.append({
                "name": shortcut_name,
                "path": shortcut_path,
                "success": success,
                "message": message
            })
            
            if success:
                success_count += 1
            else:
                failed_count += 1
                
        return (success_count, failed_count, results)

    def backup_shortcuts(self, backup_dir=None):
        """
        Create a backup of all shortcuts.
        
        Args:
            backup_dir: Directory to save backups (default: create a new one)
            
        Returns:
            (success, backup_path)
        """
        if backup_dir is None:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            backup_dir = os.path.join(os.path.expanduser("~"), f"ShortcutBackup_{timestamp}")
        
        try:
            os.makedirs(backup_dir, exist_ok=True)
            
            shortcuts = self.find_shortcuts()
            backup_count = 0
            
            for shortcut_path in shortcuts:
                shortcut_name = os.path.basename(shortcut_path)
                backup_path = os.path.join(backup_dir, shortcut_name)
                
                # Copy the shortcut file
                if sys.platform == "win32":
                    import shutil
                    shutil.copy2(shortcut_path, backup_path)
                else:
                    # In demo mode, just copy the file
                    with open(shortcut_path, "r") as src, open(backup_path, "w") as dst:
                        dst.write(src.read())
                
                backup_count += 1
                
            return (True, backup_dir, backup_count)
            
        except Exception as e:
            return (False, None, f"Error creating backup: {e}")

    def restore_shortcuts(self, backup_dir, location="user"):
        """
        Restore shortcuts from a backup.
        
        Args:
            backup_dir: Directory containing the backup
            location: Where to restore the shortcuts ("user" or "common")
            
        Returns:
            (success_count, failed_count, results)
        """
        if not os.path.exists(backup_dir):
            return (0, 0, [{"success": False, "message": "Backup directory not found"}])
            
        dest_dir = self.user_start_menu if location == "user" else self.common_start_menu
        
        if location == "common" and sys.platform == "win32" and not self.is_admin():
            return (0, 0, [{"success": False, "message": "Admin privileges required to restore to All Users"}])
            
        success_count = 0
        failed_count = 0
        results = []
        
        try:
            os.makedirs(dest_dir, exist_ok=True)
            
            for root, dirs, files in os.walk(backup_dir):
                for file in files:
                    if file.lower().endswith(".lnk"):
                        src_path = os.path.join(root, file)
                        
                        # Calculate the relative path from backup_dir
                        rel_path = os.path.relpath(root, backup_dir)
                        
                        # Create the destination directory
                        dest_path = dest_dir if rel_path == "." else os.path.join(dest_dir, rel_path)
                        os.makedirs(dest_path, exist_ok=True)
                        
                        # Full path to the destination file
                        dst_path = os.path.join(dest_path, file)
                        
                        try:
                            # Copy the shortcut file
                            if sys.platform == "win32":
                                import shutil
                                shutil.copy2(src_path, dst_path)
                            else:
                                # In demo mode, just copy the file
                                with open(src_path, "r") as src, open(dst_path, "w") as dst:
                                    dst.write(src.read())
                                    
                            success_count += 1
                            results.append({
                                "name": file,
                                "success": True,
                                "message": f"Restored to {dst_path}"
                            })
                        except Exception as e:
                            failed_count += 1
                            results.append({
                                "name": file,
                                "success": False,
                                "message": f"Failed to restore: {e}"
                            })
                            
            return (success_count, failed_count, results)
            
        except Exception as e:
            return (0, 0, [{"success": False, "message": f"Error during restoration: {e}"}])


def main():
    """Main function for standalone testing."""
    print("Start Menu Shortcut Verifier and Repair Tool")
    print("===========================================")
    
    verifier = ShortcutVerifier()
    
    print("\nFinding shortcuts...")
    shortcuts = verifier.find_shortcuts()
    print(f"Found {len(shortcuts)} shortcuts")
    
    print("\nVerifying shortcuts...")
    valid_count, broken_count, shortcuts_info = verifier.verify_all_shortcuts()
    print(f"Results: {valid_count} valid, {broken_count} broken")
    
    if broken_count > 0:
        print("\nBroken shortcuts:")
        for info in shortcuts_info:
            if not info["valid"]:
                print(f"- {info['name']} -> {info['target']} ({info['error']})")
        
        print("\nAttempting to repair broken shortcuts...")
        success_count, failed_count, repair_results = verifier.repair_all_shortcuts()
        print(f"Repair results: {success_count} fixed, {failed_count} failed")
        
        if success_count > 0:
            print("\nRepaired shortcuts:")
            for result in repair_results:
                if result["success"]:
                    print(f"- {result['name']} - {result['message']}")
        
        if failed_count > 0:
            print("\nShortcuts that could not be repaired:")
            for result in repair_results:
                if not result["success"]:
                    print(f"- {result['name']} - {result['message']}")
    
    print("\nCreating backup of all shortcuts...")
    success, backup_path, backup_count = verifier.backup_shortcuts()
    if success:
        print(f"Backup created successfully at {backup_path}")
        print(f"Backed up {backup_count} shortcuts")
    else:
        print(f"Backup failed: {backup_path}")

if __name__ == "__main__":
    main()