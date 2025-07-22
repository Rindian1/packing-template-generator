# Packing Template Generator - Technical Brief

## Overview
Create a Python-based desktop application that generates packing templates (ITP and Packing Slip) as XLSX files. The application should provide a user-friendly GUI for selecting template types, inputting equipment details, and exporting the data to a formatted XLSX file.

## Core Features

### 1. Template Management
- Support multiple template types: "ITP" and "Packing Slip"
- Each template type should have predefined equipment options:
  - "33kV Disc Manual + Single E/S Manual"
  - "72.5kV Disc Motorised + Dual E/S Motorised"

### 2. User Interface
- Clean, intuitive GUI using Tkinter
- Form fields for the following information:
  - Equipment (dropdown selection)
  - Customer (text input)
  - Purchase Order (text input)
  - Date (auto-filled with current date if empty)
  - Drawing No. (text input)
  - Serial No. (text input)
- Export button to save the template

### 3. Data Processing
- Collect user input from form fields
- Auto-fill current date if date field is empty
- Validate required fields before export

### 4. File Operations
- Save dialog for selecting export location
- Generate XLSX file based on selected template
- Preserve template formatting while inserting user data

## Technical Requirements

### 1. Dependencies
- Python 3.x
- Tkinter 
- shutil module 
- datetime module 
- openpyxl module 

### 2. File Structure
```
project/
├── main.py              # Main application file
├── PACKING LIST - Template.xls  # Template xls file
├── README.md            # Project documentation 
├─ ITP LIST -Template.xls

```

### 3. XLS Template Format
- The application should work with a template XLS file
- User data should be inserted into specific cells in the template
- First column contains field labels, second column contains values
- Data should be inserted starting from row 3 - 8, on column C

## Implementation Notes
- The application should be single-window for simplicity
- Error handling for file operations is required
- The UI should be responsive and provide feedback on operations
- Default values should be provided where possible
- The export should maintain the structure of the template XLS

## Expected Behavior
1. User selects template type (ITP or Packing Slip)
2. User selects equipment from available options
3. User fills in the remaining fields
4. On export, application:
   - Validates required fields
   - Prompts for save location
   - Creates a copy of the template
   - Inserts user data into the appropriate cells
   - Saves the file
   - Shows success/error message 
!!NOTE: THE TEMPLATES FORMATTING MUST BE PRESERVED, THIS INCLUDES CELL SIZES, COLORS, AND ANY OTHER FORMATTING!! 

### DEPLOYMENT 
- The application should be deployed as a single executable file
- The application should be able to run on any Windows machine

## Error Handling
- Handle missing template file
- Validate required fields before export
- Handle file write permissions
- Provide meaningful error messages to the user

