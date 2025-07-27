# Excel Consolidator

A GUI application that allows users to consolidate multiple sheets from an Excel file into a single sheet, with the ability to specify the header row.

## Features

- User-friendly graphical interface
- Select Excel files (.xlsx or .xls)
- Specify the header row number
- Consolidate all sheets into a single sheet
- Automatic naming of output file

## Installation

### From Source

1. Clone this repository
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the application:
   ```
   python excel_consolidator_gui.py
   ```

### Building the Executable

To create a standalone executable:

1. Install PyInstaller if you haven't already:
   ```
   pip install pyinstaller
   ```

2. Build the executable:
   ```
   pyinstaller --onefile --windowed excel_consolidator_gui.py
   ```

3. The executable will be created in the `dist` folder

## Usage

1. Launch the application
2. Click "Select Excel File" to choose your Excel file
3. Use the spin box to specify which row contains the headers (0-based index)
4. Click "Consolidate" to process the file
5. The consolidated file will be saved in the same directory as the input file with "_consolidated" appended to the filename

## Requirements

- Python 3.8 or higher
- Windows operating system
- Required Python packages (see requirements.txt) 