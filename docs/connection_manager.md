# E3 Connection Manager

This module provides functionality to connect to E3.series applications with support for selecting from multiple running instances, similar to the VBA `ConnectToE3()` function.

## Overview

The E3 Connection Manager automatically detects running E3.series instances and provides a user-friendly interface for selecting which instance to connect to when multiple instances are running.

### Key Features

- **Automatic Instance Detection**: Detects all running E3.series processes
- **Multiple Instance Support**: Shows a selection dialog when multiple instances are detected
- **Single Instance Auto-Connect**: Automatically connects when only one instance is running
- **Error Handling**: Provides clear error messages and user guidance
- **Integration Ready**: Easy to integrate into existing E3 automation scripts

## Usage

### Method 1: Using the Connection Manager Class

```python
from lib.e3_connection_manager import E3ConnectionManager

# Create connection manager
manager = E3ConnectionManager(logger)

# Connect to E3 (will show selector if multiple instances)
app = manager.connect_to_e3()

if app:
    # Use the E3 application object
    job = app.CreateJobObject()
    # ... rest of your E3 automation code
```

### Method 2: Using the Convenience Function

```python
from lib.e3_connection_manager import create_e3_connection

# Simple one-line connection
app = create_e3_connection(logger)

if app:
    # Use the E3 application object
    job = app.CreateJobObject()
    # ... rest of your E3 automation code
```

### Method 3: Updated E3 Scripts

All existing E3 automation scripts have been updated to use the connection manager:

- `lib/e3_wire_numbering.py`
- `lib/e3_device_designation.py` 
- `lib/e3_terminal_pin_names.py`

These scripts now automatically handle multiple E3 instances.

## How It Works

### Single Instance
When only one E3.series instance is running:
- Connects automatically without user interaction
- Logs the successful connection

### Multiple Instances
When multiple E3.series instances are detected:
- Shows a selection dialog listing all instances
- User selects which instance to connect to
- Connects to the selected instance

### No Instances
When no E3.series instances are found:
- Shows an error dialog
- Provides guidance to start E3.series with a project open

## Selection Dialog

The selection dialog shows:
- Process ID (PID) for each E3.series instance
- Easy-to-use list selection interface
- Connect and Cancel buttons
- Keyboard shortcuts (Enter to connect, Escape to cancel)

## Integration with Existing Scripts

The connection manager has been integrated into all existing E3 automation scripts. The `connect_to_e3()` method in each script class now uses the connection manager automatically.

### Before (Old Method)
```python
def connect_to_e3(self):
    try:
        pythoncom.CoInitialize()
        self.app = e3series.Application()
        # ... rest of connection code
```

### After (New Method)
```python
def connect_to_e3(self):
    try:
        pythoncom.CoInitialize()
        connection_manager = E3ConnectionManager(self.logger)
        self.app = connection_manager.connect_to_e3()
        
        if not self.app:
            return False
        # ... rest of connection code
```

## Error Handling

The connection manager provides comprehensive error handling:

- **COM Initialization**: Handles COM initialization errors
- **Process Detection**: Gracefully handles process enumeration failures
- **Connection Failures**: Provides clear error messages
- **User Cancellation**: Handles user cancelling the selection dialog

## Testing

Test scripts are provided to verify functionality:

- `apps/test_e3_connection_manager.py` - Full connection test with user interaction
- `apps/test_e3_detection.py` - Instance detection test without user interaction
- `apps/debug_e3_instances.py` - Detailed diagnostic information

## Requirements

- Python 3.8+
- `e3series` package
- `win32com.client` (pywin32)
- `tkinter` (usually included with Python)
- Running E3.series instance(s)

## Compatibility

This implementation is based on the VBA `ConnectToE3()` function and provides similar functionality:

- Detects multiple E3.series instances
- Provides user selection interface
- Handles single instance auto-connection
- Provides error handling and user guidance

The main difference is that this Python implementation uses process enumeration instead of CT.Dispatcher due to API limitations, but provides the same user experience.
