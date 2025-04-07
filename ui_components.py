import os
import sys
import platform
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QFileDialog, QComboBox, QCheckBox, QFrame, QMessageBox, QGroupBox,
    QFormLayout, QRadioButton, QButtonGroup, QSizePolicy, QSpacerItem,
    QApplication
)
from PyQt5.QtCore import Qt, QMimeData, QUrl, pyqtSignal, QSize
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QIcon, QPalette, QColor, QFont, QFontMetrics
from PyQt5.QtSvg import QSvgWidget
from styles import StyleSheet

# Check if running on Windows
IS_WINDOWS = platform.system() == "Windows"

class DropArea(QLabel):
    """Custom widget for drag and drop file selection."""
    fileDropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedHeight(200)
        self.setStyleSheet(StyleSheet.DROP_AREA)
        
        # Create layout
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        # Add SVG icon
        self.svg_widget = QSvgWidget("assets/drag_drop.svg")
        self.svg_widget.setFixedSize(64, 64)
        layout.addWidget(self.svg_widget, alignment=Qt.AlignCenter)
        
        # Add text labels
        self.title_label = QLabel("Drag and drop .exe file here")
        self.title_label.setStyleSheet(StyleSheet.DROP_AREA_TITLE)
        layout.addWidget(self.title_label, alignment=Qt.AlignCenter)
        
        self.subtitle_label = QLabel("or click to browse")
        self.subtitle_label.setStyleSheet(StyleSheet.DROP_AREA_SUBTITLE)
        layout.addWidget(self.subtitle_label, alignment=Qt.AlignCenter)
        
        self.setLayout(layout)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter events."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(StyleSheet.DROP_AREA_DRAG_OVER)
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """Handle drag leave events."""
        self.setStyleSheet(StyleSheet.DROP_AREA)
        
    def dropEvent(self, event: QDropEvent):
        """Handle drop events."""
        urls = event.mimeData().urls()
        if urls and len(urls) > 0:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.exe'):
                self.fileDropped.emit(file_path)
            else:
                QMessageBox.warning(self, "Invalid File", "Please select a valid .exe file.")
        
        self.setStyleSheet(StyleSheet.DROP_AREA)
        
    def mousePressEvent(self, event):
        """Handle mouse click events."""
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(
            self, "Select Executable", "", "Executable Files (*.exe);;All Files (*)"
        )
        
        if file_path:
            self.fileDropped.emit(file_path)

class PreviewWidget(QWidget):
    """Widget to preview shortcut information before creation."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(StyleSheet.PREVIEW_WIDGET)
        
        # Main layout
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("Shortcut Preview")
        title_label.setStyleSheet(StyleSheet.SECTION_TITLE)
        layout.addWidget(title_label)
        
        # Preview content frame
        self.preview_frame = QFrame()
        self.preview_frame.setStyleSheet(StyleSheet.PREVIEW_FRAME)
        preview_layout = QFormLayout(self.preview_frame)
        
        # Preview fields
        self.executable_path = QLabel("No file selected")
        self.executable_path.setWordWrap(True)
        
        self.shortcut_name = QLabel("")
        self.destination = QLabel("")
        self.product_info = QLabel("")
        self.company_info = QLabel("")
        
        # Add fields to form layout
        preview_layout.addRow("Executable:", self.executable_path)
        preview_layout.addRow("Shortcut Name:", self.shortcut_name)
        preview_layout.addRow("Destination:", self.destination)
        preview_layout.addRow("Product:", self.product_info)
        preview_layout.addRow("Publisher:", self.company_info)
        
        layout.addWidget(self.preview_frame)
        self.setLayout(layout)
        
    def update_preview(self, exe_info, shortcut_name, destination):
        """Update the preview with executable and shortcut information."""
        self.executable_path.setText(exe_info.get('path', 'Unknown path'))
        self.shortcut_name.setText(shortcut_name)
        self.destination.setText(destination)
        self.product_info.setText(exe_info.get('product_name', 'Unknown') + 
                                  (f" (v{exe_info.get('version', '')})" if exe_info.get('version') else ""))
        self.company_info.setText(exe_info.get('company', 'Unknown publisher'))

class ShortcutCreatorUI(QWidget):
    """Main UI for the shortcut creator application."""
    def __init__(self, parent=None, shortcut_creator=None):
        super().__init__(parent)
        self.shortcut_creator = shortcut_creator
        self.current_exe_path = None
        self.exe_info = {}
        self.init_ui()
        
    def init_ui(self):
        """Initialize the UI components."""
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # Title
        title_label = QLabel("Start Menu Shortcut Creator")
        title_label.setStyleSheet(StyleSheet.APP_TITLE)
        main_layout.addWidget(title_label, alignment=Qt.AlignCenter)
        
        # Description
        desc_label = QLabel("Create shortcuts in the Windows Start Menu for your applications")
        desc_label.setStyleSheet(StyleSheet.APP_DESCRIPTION)
        main_layout.addWidget(desc_label, alignment=Qt.AlignCenter)
        
        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet(StyleSheet.SEPARATOR)
        main_layout.addWidget(separator)
        
        # File selection section
        file_section_label = QLabel("1. Select Application")
        file_section_label.setStyleSheet(StyleSheet.SECTION_TITLE)
        main_layout.addWidget(file_section_label)
        
        # Drop area
        self.drop_area = DropArea()
        self.drop_area.fileDropped.connect(self.handle_file_selection)
        main_layout.addWidget(self.drop_area)
        
        # Shortcut settings section
        settings_section_label = QLabel("2. Configure Shortcut")
        settings_section_label.setStyleSheet(StyleSheet.SECTION_TITLE)
        main_layout.addWidget(settings_section_label)
        
        # Settings content
        settings_group = QGroupBox()
        settings_group.setStyleSheet(StyleSheet.SETTINGS_GROUP)
        settings_layout = QFormLayout(settings_group)
        
        # Shortcut name
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Enter shortcut name")
        self.name_input.setEnabled(False)
        self.name_input.textChanged.connect(self.update_preview)
        settings_layout.addRow("Shortcut Name:", self.name_input)
        
        # Destination options
        destination_widget = QWidget()
        destination_layout = QVBoxLayout(destination_widget)
        destination_layout.setContentsMargins(0, 0, 0, 0)
        
        self.destination_group = QButtonGroup(self)
        
        self.current_user_radio = QRadioButton("Current User (no admin rights needed)")
        self.current_user_radio.setChecked(True)
        self.current_user_radio.toggled.connect(self.update_preview)
        
        self.all_users_radio = QRadioButton("All Users (requires admin privileges)")
        self.all_users_radio.toggled.connect(self.update_preview)
        
        self.destination_group.addButton(self.current_user_radio)
        self.destination_group.addButton(self.all_users_radio)
        
        destination_layout.addWidget(self.current_user_radio)
        destination_layout.addWidget(self.all_users_radio)
        
        settings_layout.addRow("Install For:", destination_widget)
        
        # Subfolder option
        self.subfolder_input = QLineEdit()
        self.subfolder_input.setPlaceholderText("Optional: create in subfolder")
        self.subfolder_input.setEnabled(False)
        self.subfolder_input.textChanged.connect(self.update_preview)
        settings_layout.addRow("Subfolder:", self.subfolder_input)
        
        main_layout.addWidget(settings_group)
        
        # Preview section
        preview_section_label = QLabel("3. Preview")
        preview_section_label.setStyleSheet(StyleSheet.SECTION_TITLE)
        main_layout.addWidget(preview_section_label)
        
        self.preview_widget = PreviewWidget()
        main_layout.addWidget(self.preview_widget)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        self.create_button = QPushButton("Create Shortcut")
        self.create_button.setStyleSheet(StyleSheet.PRIMARY_BUTTON)
        self.create_button.setEnabled(False)
        self.create_button.clicked.connect(self.create_shortcut)
        
        self.cancel_button = QPushButton("Reset")
        self.cancel_button.setStyleSheet(StyleSheet.SECONDARY_BUTTON)
        self.cancel_button.clicked.connect(self.reset_form)
        
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.create_button)
        
        main_layout.addLayout(button_layout)
        
        # Add stretch to push everything to the top
        main_layout.addStretch()
        
        self.setLayout(main_layout)
    
    def handle_file_selection(self, file_path):
        """Process selected executable file."""
        # Validate the file
        if not self.shortcut_creator.is_valid_exe(file_path):
            QMessageBox.warning(
                self, 
                "Invalid File", 
                "The selected file is not a valid Windows executable."
            )
            return
        
        # Store file path and fetch info
        self.current_exe_path = file_path
        self.exe_info = self.shortcut_creator.get_exe_info(file_path)
        self.exe_info['path'] = file_path
        
        # Update UI
        self.name_input.setEnabled(True)
        self.subfolder_input.setEnabled(True)
        self.create_button.setEnabled(True)
        
        # Pre-populate the shortcut name with suggested name
        suggested_name = self.exe_info.get('suggested_name', os.path.splitext(os.path.basename(file_path))[0])
        self.name_input.setText(suggested_name)
        
        # Update preview
        self.update_preview()
    
    def update_preview(self):
        """Update the shortcut preview."""
        if not self.current_exe_path:
            return
        
        shortcut_name = self.name_input.text() or "Unnamed Shortcut"
        
        # Determine destination text
        if self.all_users_radio.isChecked():
            destination_base = "All Users Start Menu"
            destination_path = self.shortcut_creator.common_start_menu
        else:
            destination_base = "Current User Start Menu"
            destination_path = self.shortcut_creator.user_start_menu
        
        # Add subfolder if specified
        subfolder = self.subfolder_input.text()
        if subfolder:
            destination = f"{destination_base} → {subfolder}"
            destination_path = os.path.join(destination_path, subfolder)
        else:
            destination = destination_base
        
        # Show full path on hover
        destination_with_path = f"{destination}\n({destination_path})"
        
        # Update preview widget
        self.preview_widget.update_preview(
            self.exe_info, 
            shortcut_name + ".lnk", 
            destination_with_path
        )
    
    def create_shortcut(self):
        """Create the shortcut in the Start Menu."""
        if not self.current_exe_path:
            return
        
        shortcut_name = self.name_input.text()
        if not shortcut_name:
            QMessageBox.warning(self, "Missing Information", "Please enter a name for the shortcut.")
            return
        
        for_all_users = self.all_users_radio.isChecked()
        
        # Check if we need admin rights and prompt if necessary
        if for_all_users and not self.shortcut_creator.is_admin():
            result = QMessageBox.question(
                self,
                "Administrator Rights Required",
                "Creating shortcuts for all users requires administrator privileges. "
                "Do you want to restart the application with elevated privileges?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if result == QMessageBox.Yes:
                # Restart with admin rights - only works on Windows
                if IS_WINDOWS:
                    import ctypes
                    ctypes.windll.shell32.ShellExecuteW(
                        None, 
                        "runas", 
                        sys.executable, 
                        " ".join(sys.argv), 
                        None, 
                        1
                    )
                else:
                    QMessageBox.information(
                        self,
                        "Demo Mode",
                        "In demo mode, administrator elevation is simulated. "
                        "This would attempt to restart the application with admin rights on a real Windows system."
                    )
                QApplication.quit()
                return
            else:
                # User declined admin elevation, switch to current user
                self.current_user_radio.setChecked(True)
                for_all_users = False
        
        # Create the shortcut
        subfolder = self.subfolder_input.text() if self.subfolder_input.text() else None
        success, message = self.shortcut_creator.create_shortcut(
            self.current_exe_path,
            shortcut_name,
            for_all_users,
            subfolder
        )
        
        # Show result message
        if success:
            QMessageBox.information(self, "Success", message)
            self.reset_form()
        else:
            QMessageBox.critical(self, "Error", message)
    
    def reset_form(self):
        """Reset the form to initial state."""
        self.current_exe_path = None
        self.exe_info = {}
        
        # Reset UI elements
        self.name_input.clear()
        self.name_input.setEnabled(False)
        self.subfolder_input.clear()
        self.subfolder_input.setEnabled(False)
        self.current_user_radio.setChecked(True)
        self.create_button.setEnabled(False)
        
        # Reset preview
        self.preview_widget.executable_path.setText("No file selected")
        self.preview_widget.shortcut_name.setText("")
        self.preview_widget.destination.setText("")
        self.preview_widget.product_info.setText("")
        self.preview_widget.company_info.setText("")
