# Attendance Management System

A modern, feature-rich attendance tracking application developed by Ahmad Tchnology. This application allows for real-time attendance tracking using NFC/ID cards, with filters for different classes and groups.

![Attendance System Screenshot](screenshot.png)

## Features

- **Modern UI Design**: Clean, professional interface with tabbed navigation and consistent styling
- **Student Management**: Add, import, and manage student records
- **Real-time Attendance Tracking**: Record attendance with NFC/ID cards
- **Smart Filtering**: Automatically filters by Major, Stage, and Group based on the first student
- **Export Functionality**: Export attendance data to Excel files
- **Database Management**: Create and select different database files for different classes or courses
- **Cross-compatible Study Types**: Support for Morning, Evening, and Hosted student types

## Key Improvements

### UI Enhancements
- Modern color scheme with professional blues and accent colors
- Tabbed interface for better organization between student management and attendance recording
- Rounded corners for panels with subtle borders
- Proper form layouts with consistent styling
- Improved treeview for attendance dashboard
- Better visual hierarchy with primary and secondary elements

### Functional Improvements
- Automatic database creation and management
- Improved error handling for missing images and resources
- Better form validation and feedback
- Enhanced Excel export with automatic naming based on filters
- Consistent use of geometry managers (pack) to avoid layout conflicts

## Requirements

- Python 3.9 or higher
- Required packages:
  - tkinter
  - sqlite3
  - openpyxl
  - Pillow (PIL)

## Installation

### From Source
1. Clone this repository
2. Install the required packages: `pip install -r requirements.txt`
3. Run the application: `python app.py`

### Pre-built Executable
1. Download the latest `Attendance_System_Ahmad_Tchnology.exe` from the releases page
2. Run the executable - no installation required!

## Building the Executable

To build the executable yourself:

1. Ensure you have PyInstaller installed: `pip install pyinstaller`
2. Place your logo images in the project directory as `Any_logo.png` and `Any_logo2.png`
3. Run: `pyinstaller attendance_system.spec`
4. The executable will be created in the `dist` folder

## Usage

See the included `HOW_TO_USE.txt` file for detailed instructions on how to use the application.

## Data Structure

The application uses SQLite databases with the following structure:

### Students Table
- student_id (Primary Key)
- name
- major
- stage
- study (Morning, Evening, or Hosted)
- group_name

### Attendance Table
- id (Primary Key, Autoincrement)
- student_id (Foreign Key)
- name
- major
- stage
- study
- group_name
- timestamp
- attended (1 = present)

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

Developed by Ahmad Tchnology - an advanced attendance management solution that combines ease of use with powerful features.