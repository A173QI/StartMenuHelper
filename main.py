import sys
import os
import platform
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

# Use demo implementation for non-Windows environments
from shortcut_creator_demo import ShortcutCreator
from ui_components import ShortcutCreatorUI

def is_admin():
    """Check if the current process has admin privileges."""
    # This is a demo implementation that always returns False on non-Windows platforms
    return False

def run_as_admin():
    """Re-run the program with admin privileges (demo version)."""
    # This is just a placeholder for demo purposes
    pass

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.shortcut_creator = ShortcutCreator()
        self.ui = ShortcutCreatorUI(self, self.shortcut_creator)
        self.init_ui()
        
    def init_ui(self):
        self.setCentralWidget(self.ui)
        self.setWindowTitle("Start Menu Shortcut Creator")
        self.setMinimumSize(800, 600)
        # Set app icon (will be shown as SVG inline in our app)
        self.setWindowIcon(QIcon("assets/app_icon.svg"))
        self.center_on_screen()

    def center_on_screen(self):
        """Center the window on the screen."""
        screen_geometry = QApplication.desktop().availableGeometry()
        window_geometry = self.frameGeometry()
        center_point = screen_geometry.center()
        window_geometry.moveCenter(center_point)
        self.move(window_geometry.topLeft())

def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    
    # Set application style to match Windows
    app.setStyle("WindowsVista")
    
    # Set up high DPI scaling
    app.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
