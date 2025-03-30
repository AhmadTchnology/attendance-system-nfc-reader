# Attendance System Using NFC Reader

A desktop application for managing student attendance using NFC card scanning technology.

## Overview

This application provides an easy-to-use interface for tracking student attendance at educational institutions. It allows administrators to create databases for different classes, import student information, record attendance using NFC cards, and export attendance data for reporting purposes.

## Features

- **Real-time Attendance Tracking**: Scan NFC cards to instantly mark students as present
- **Database Management**: Create and manage multiple class databases
- **Student Management**: Add students individually or import from Excel files
- **Grouping and Filtering**: Automatically filters students by major, stage, study type, and group
- **Attendance Reports**: Export attendance data to Excel for further analysis
- **User-friendly Interface**: Modern UI with intuitive controls

## Requirements

- Python 3.6+
- Required Python packages:
  - tkinter (for the GUI)
  - sqlite3 (for database management)
  - openpyxl (for Excel file handling)
  - PIL/Pillow (for image processing)

## Installation

1. Clone this repository
2. Install the required dependencies:
   ```
   pip install openpyxl pillow
   ```
3. Run the application:
   ```
   python app.py
   ```

## Usage

### Setting up a Database

1. Launch the application
2. Click "Create New Database" to create a new class database or select an existing one from the dropdown
3. The database will be stored in the "Students Databases" folder

### Adding Students

#### Manually:
1. Fill in the student information (ID, Name, Major, Stage, Study, Group)
2. Click "Add Student"

#### Import from Excel:
1. Click "Import Students from Excel"
2. Select an Excel file with student information
3. The Excel file should have columns for: Student ID, Name, Major, Stage, Study, Group

### Recording Attendance

1. Make sure the correct database is selected
2. Have students scan their NFC cards or manually enter their ID in the "Scan NFC Card" field
3. The first student who scans sets the filters (major, stage, study, group) for that session
4. Attendance is recorded in real-time and displayed in the dashboard

### Exporting Attendance Data

1. Click "Export Attendance" to save the current attendance data
2. Choose a location to save the Excel file
3. The exported file includes all students from the filtered group with their attendance status

### Resetting Attendance

1. Click "Reset Attendance" to clear all attendance records
2. This will reset all filters and allow for a new attendance session

## Project Structure

- `app.py`: Main application file
- `Students Databases/`: Directory containing database files and Excel templates
  - `*.db`: SQLite database files for different classes
  - `*.xlsx`: Excel files containing student information
- `Stages/`: Directory containing stage-specific information
  - `First Stage/`: First year student data
  - `Second Stage/`: Second year student data
  - `Third Stage/`: Third year student data
  - `Fourth Stage/`: Fourth year student data
- Logo files: Any Image You Want

## Directory Tree

```
.
│   .gitignore
│   app.py
│   README.md
│   Any_Image_You_Want.png
│
├───Students Databases
│   │   First.db (Database file)
│   │   Any_Excel_File_You_Want.xlsx (Excel template)
│
└───Stages
    ├───First Stage
    │   ├───Evening
    │   │   ├───GA
    │   │   └───GB
    │   └───Morning
    │       ├───GA
    │       └───GB
    ├───Fourth Stage
    │   ├───Evening
    │   │   ├───GA
    │   │   └───GB
    │   └───Morning
    │       ├───GA
    │       └───GB
    ├───Second Stage
    │   ├───Evening
    │   │   ├───GA
    │   │   └───GB
    │   └───Morning
    │       ├───GA
    │       └───GB
    └───Third Stage
        ├───Evening
        │   ├───GA
        │   └───GB
        └───Morning
            ├───GA
            └───GB
```

This structure organizes student data by Stage (year), Study type (Morning/Evening), and Group (GA/GB). The application uses this organization to manage and filter attendance records , You can change the structure as you want.

## Database Schema

### Students Table
- `student_id` (TEXT): Primary key, unique identifier for each student
- `name` (TEXT): Student's full name
- `major` (TEXT): Student's major/specialization
- `stage` (TEXT): Student's year/stage
- `study` (TEXT): Study type (e.g., Morning, Evening)
- `group_name` (TEXT): Student's assigned group

### Attendance Table
- `id` (INTEGER): Primary key, auto-incremented
- `student_id` (TEXT): Foreign key to students table
- `name` (TEXT): Student's name
- `major` (TEXT): Student's major
- `stage` (TEXT): Student's stage
- `study` (TEXT): Study type
- `group_name` (TEXT): Student's group
- `timestamp` (TEXT): Date and time when attendance was recorded
- `attended` (INTEGER): Boolean flag (1 = present, 0 = absent)

## Excel File Structure

The application accepts Excel files with the following structure:
- Column 1: Student ID
- Column 2: Name
- Column 3: Major
- Column 4: Stage
- Column 5: Study
- Column 6: Group

The first row should contain headers and will be skipped during import.

## License

This project is proprietary and has been developed for educational institutions.

## Author

Made with ❤️ by [@AhmadTchnology](https://github.com/AhmadTchnology)