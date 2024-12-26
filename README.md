# SAP Automation Script

Python script for automating SAP GUI interactions using SAP GUI Scripting API and win32 library.

## Features
- SAP system login/logout automation
- Automatic password change handling
- Multiple login session management
- Element interaction with timeout handling
- Window and dialog management
- Cell value operations
- Resource cleanup

## Prerequisites
- Python 3.6+
- SAP GUI Client
- SAP GUI Scripting enabled
- Windows OS

## Installation
```bash
pip install pywin32 pygetwindow psutil
```

## Usage

```python
from sap_gui import SapGui

# Configure SAP connection
sap_args = {
    "platform": "SAP PRD",
    "client": "100",
    "username": "USER",
    "password": "PASSWORD",
    "language": "EN",
    "path": "saplogon.exe"  # Default path
}

try:
    # Initialize SAP session
    sap = SapGui(sap_args)

    # Login
    if sap.sapLogin():
        # Execute SAP operations
        sap.perform_operation("/nVA01", "wnd[0]/usr/ctxtVBAK-AUART")

        # Other operations...

    # Cleanup
    sap.sapLogout()
    sap.close_connection()

except Exception as e:
    print(f"Error: {e}")
```

## Key Functions
- `sapLogin()`: SAP system login with password change handling
- `perform_operation()`: Execute SAP transactions/commands
- `wait_for_element()`: Wait for SAP GUI elements
- `set_cell_value()`: Set values in table cells
- `bring_dialog_to_top()`: Handle SAP dialogs
- `scroll_to_field()`: Navigate to fields
- `get_sap_element_text()`: Retrieve element text

## Error Handling
The script includes comprehensive error handling and logging for:
- Connection failures
- Login issues
- Password changes
- Element interactions
- Resource cleanup
