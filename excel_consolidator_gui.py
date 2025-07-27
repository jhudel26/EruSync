import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                           QWidget, QFileDialog, QLabel, QSpinBox, QMessageBox,
                           QFrame, QHBoxLayout, QSpacerItem, QSizePolicy)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor

class EruSync(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EruSync - Excel Consolidator")
        self.setGeometry(100, 100, 700, 600) # Slightly adjusted size
        try:
            # Attempt to load icon, handle potential FileNotFoundError
            self.setWindowIcon(QIcon('eru_sync_icon.ico'))
        except FileNotFoundError:
            print("Warning: eru_sync_icon.ico not found. Using default window icon.")
            # Optionally set a default icon or leave it as system default
        
        # Set global font
        font = QFont("Segoe UI", 10) # Consistent modern font
        QApplication.setFont(font)

        # --- Modernized Global Stylesheet ---
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5; /* Keep light background */
            }
            QLabel {
                color: #333333; 
                font-size: 10pt;
            }
            QPushButton {
                background-color: #008080; /* Teal */
                color: white;
                border: none;
                padding: 10px 18px; 
                border-radius: 4px; 
                font-size: 10pt;
                font-weight: 500; 
                min-height: 28px; 
                transition: background-color 0.3s ease; 
            }
            QPushButton:hover {
                background-color: #006666; /* Darker Teal on hover */
            }
            QPushButton:disabled {
                background-color: #c0c0c0; 
                color: #f0f2f5;
            }
            QSpinBox {
                padding: 8px 10px; 
                border: 1px solid #ced4da; 
                border-radius: 4px;
                background-color: white;
                color: #333333; 
                font-size: 10pt;
                min-width: 60px; 
                max-width: 100px;
            }
            QFrame#InputFrame { 
                background-color: #ffffff; 
                border-radius: 6px; 
                padding: 18px; 
                border: 1px solid #e0e0e0; 
                margin-bottom: 15px; 
            }
            QLabel#SectionTitle {
                font-size: 13pt; 
                font-weight: 600; 
                color: #006666; /* Darker Teal title */
                margin-bottom: 12px;
                border-bottom: 1px solid #e0e0e0; 
                padding-bottom: 5px;
            }
            QLabel#InfoLabel {
                font-size: 9pt; 
                color: #555555; 
                padding: 10px;
                background-color: #e9ecef; 
                border-radius: 4px;
                border: 1px solid #d6d9dc; 
            }
            QStatusBar {
                background-color: #e9ecef;
                color: #333333;
                font-size: 9pt;
                padding: 3px 5px;
            }
            QToolTip { 
                background-color: #333333;
                color: white;
                border: 1px solid #333333;
                padding: 5px;
                border-radius: 3px;
            }
        """)
        
        # --- Central Widget and Main Layout ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0) # Let frames handle margin-bottom
        main_layout.setContentsMargins(30, 30, 30, 30) # Adjusted margins
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        # --- Title ---
        title_label = QLabel("EruSync")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("""
            font-size: 22pt; 
            font-weight: bold;
            color: #008080; /* Teal title */
            margin-bottom: 5px; 
        """)
        main_layout.addWidget(title_label)
        
        # --- Description ---
        desc_label = QLabel("Excel Sheet Consolidation Tool")
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setStyleSheet("""
            font-size: 11pt;
            color: #6c757d; 
            margin-bottom: 25px; /* Adjusted margin */
        """)
        main_layout.addWidget(desc_label)
        
        # --- File Selection Frame ---
        file_frame = QFrame()
        file_frame.setObjectName("InputFrame") 
        file_layout = QVBoxLayout(file_frame)
        file_layout.setSpacing(12) # Spacing inside frame
        file_layout.setContentsMargins(15, 15, 15, 15) # Padding inside frame
        
        file_title = QLabel("1. Select Excel File")
        file_title.setObjectName("SectionTitle") 
        
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("InfoLabel") 
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setWordWrap(True) # Allow text wrapping
        
        self.select_button = QPushButton("📂 Browse...") # More standard text
        self.select_button.clicked.connect(self.select_file)
        self.select_button.setIconSize(QSize(16, 16)) # Standard icon size
        self.select_button.setToolTip("Select the Excel workbook (.xlsx, .xls) to consolidate.") # Add tooltip

        # Layout for file info and button
        file_bottom_layout = QHBoxLayout()
        file_bottom_layout.addWidget(self.file_label, 1) # Label takes available space
        file_bottom_layout.addWidget(self.select_button)

        file_layout.addWidget(file_title)
        file_layout.addLayout(file_bottom_layout)
        main_layout.addWidget(file_frame)
        
        # --- Header Selection Frame ---
        header_frame = QFrame()
        header_frame.setObjectName("InputFrame") 
        header_layout = QVBoxLayout(header_frame)
        header_layout.setSpacing(12) 
        header_layout.setContentsMargins(15, 15, 15, 15)
        
        header_title = QLabel("2. Set Header Row")
        header_title.setObjectName("SectionTitle") 
        
        header_desc_label = QLabel("Specify the row number containing your column headers.")
        header_desc_label.setStyleSheet("font-size: 9pt; color: #555555;")
        
        spin_layout = QHBoxLayout()
        header_prompt_label = QLabel("Header row:") # Clearer prompt
        
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setRange(1, 100)
        self.header_row_spin.setValue(1)
        self.header_row_spin.setToolTip("Enter the row number (starting from 1) where the data headers begin.") # Add tooltip
        self.header_row_spin.setAlignment(Qt.AlignmentFlag.AlignCenter) # Center text in spinbox

        spin_layout.addWidget(header_prompt_label) 
        spin_layout.addStretch() # Pushes spinbox to the right
        spin_layout.addWidget(self.header_row_spin)
        
        header_layout.addWidget(header_title)
        header_layout.addWidget(header_desc_label) # Added descriptive label
        header_layout.addLayout(spin_layout)
        main_layout.addWidget(header_frame)
        
        # --- Consolidate Button Area ---
        # Use a QHBoxLayout to center the button horizontally
        button_layout = QHBoxLayout()
        button_layout.addStretch() # Spacer on the left
        self.consolidate_button = QPushButton("🔄 Consolidate Sheets")
        self.consolidate_button.clicked.connect(self.consolidate_excel)
        self.consolidate_button.setEnabled(False)
        self.consolidate_button.setMinimumWidth(200) # Ensure button isn't too small
        self.consolidate_button.setStyleSheet("""
            QPushButton {
                min-height: 35px; 
                font-size: 11pt; 
                font-weight: bold;
                margin-top: 10px; 
            }
        """) 
        self.consolidate_button.setToolTip("Combine all sheets from the selected file into one.") # Add tooltip
        button_layout.addWidget(self.consolidate_button)
        button_layout.addStretch() # Spacer on the right
        main_layout.addLayout(button_layout) # Add the button layout

        # Add a vertical spacer at the bottom to push content up
        main_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        # --- Status Bar ---
        self.statusBar().showMessage("Ready")
        
        self.file_path = None
        
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path = file_path
            # Improved feedback on selection
            base_name = os.path.basename(file_path)
            max_len = 50 # Limit displayed filename length
            display_name = base_name if len(base_name) <= max_len else base_name[:max_len-3] + "..."
            self.file_label.setText(f"Selected: {display_name}")
            self.file_label.setToolTip(f"Full path: {file_path}") # Show full path in tooltip
            self.consolidate_button.setEnabled(True)
            self.statusBar().showMessage(f"File selected: {base_name}")
    
    def consolidate_excel(self):
        if not self.file_path:
            return
            
        try:
            self.statusBar().showMessage("Processing...")
            self.consolidate_button.setEnabled(False)
            self.select_button.setEnabled(False) # Disable select while processing
            QApplication.processEvents()  # Update UI
            
            # Convert 1-based row number to 0-based for pandas
            header_row = self.header_row_spin.value() - 1
            
            # Read all sheets into a dictionary
            all_sheets = pd.read_excel(
                self.file_path, 
                sheet_name=None,
                header=header_row
            )

            # Combine them with a 'Sheet Name' column
            combined_data = []
            for sheet_name, df in all_sheets.items():
                df['Sheet Name'] = sheet_name
                combined_data.append(df)

            # Concatenate all dataframes
            final_df = pd.concat(combined_data, ignore_index=True)

            # Create output filename in the same directory as input
            base_name = os.path.splitext(os.path.basename(self.file_path))[0]
            output_dir = os.path.dirname(self.file_path)
            output_file = os.path.join(output_dir, f"{base_name}_consolidated.xlsx")
            
            # Save the consolidated file
            final_df.to_excel(output_file, index=False)
            
            self.statusBar().showMessage("Processing completed successfully!")
            self.select_button.setEnabled(True) # Re-enable select button
            # Don't re-enable consolidate here, user needs to select a new file implicitly
            QMessageBox.information(
                self,
                "Success",
                f"✅ Consolidated file saved as:\\n{output_file}"
            )
            
        except PermissionError:
            self.statusBar().showMessage("Error: Permission denied.")
            QMessageBox.critical(
                self,
                "Error",
                "❌ Permission denied. Please make sure:\\n"
                "1. The Excel file is not open in another program.\\n"
                "2. You have write permissions in the target directory.\\n"
                "3. The file is not read-only."
            )
        except Exception as e:
            self.statusBar().showMessage(f"Error: {str(e)[:50]}...") 
            self.consolidate_button.setEnabled(bool(self.file_path)) # Re-enable consolidate if file still selected
            self.select_button.setEnabled(True) # Re-enable select button
            QMessageBox.critical(
                self,
                "Error",
                f"❌ An error occurred: {str(e)}\\n\\nPlease make sure:\\n"
                "1. The Excel file is valid and not corrupted.\\n"
                "2. The header row number ({self.header_row_spin.value()}) is correct for all sheets.\\n"
                "3. The file is not password-protected."
            )

def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('eru_sync_icon.ico'))  # Set icon globally for the application
    window = EruSync()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main() 