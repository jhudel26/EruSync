EruSync
A GUI application that allows users to consolidate multiple sheets from an Excel file into a single sheet, with the ability to specify the header row.

Features
User-friendly graphical interface

Select Excel files (.xlsx or .xls)

Specify the header row number

Consolidate all sheets into a single sheet

Automatic naming of output file

Installation
From Source
Clone this repository

Install the required dependencies:

bash
Copy
Edit
pip install -r requirements.txt
Run the application:

bash
Copy
Edit
python excel_consolidator_gui.py
Building the Executable
To create a standalone executable:

Install PyInstaller if you haven't already:

bash
Copy
Edit
pip install pyinstaller
Build the executable:

bash
Copy
Edit
pyinstaller --onefile --windowed excel_consolidator_gui.py
The executable will be created in the dist folder

Usage
Launch the application

Click "Select Excel File" to choose your Excel file

Use the spin box to specify which row contains the headers (0-based index)

Click "Consolidate" to process the file

The consolidated file will be saved in the same directory as the input file with _consolidated appended to the filename

Requirements
Python 3.8 or higher

Windows operating system

Required Python packages (see requirements.txt)

Releases
You can download the latest precompiled executable from the Releases section on GitHub.

No Python installation is required—just download and run the .exe file.
