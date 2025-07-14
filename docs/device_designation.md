# E3 Device Designation Automation

This script automatically updates device designations in E3.series projects based on the sheet and grid position of the topmost leftmost pin of the first symbol of each device.

**Special handling for terminal devices:** Terminal devices (identified using E3 API methods `IsTerminal()` and `IsTerminalBlock()`) are completely skipped by this script. Their designations remain unchanged. The script includes fallback detection using letter codes (T, TB, X, XT, TERM) if the API methods fail.

## Overview

The device designation automation script follows similar logic to the wire numbering script but works with devices and their symbols instead of connections and net segments.

### Designation Format

The script generates device designations using the format:
```
{device letter code}{sheet}{grid}
```

For example:
- `M1A1` - Motor (M) on sheet 1, grid A1
- `K2B3` - Contactor (K) on sheet 2, grid B3
- `T1C5` - Transformer (T) on sheet 1, grid C5

### Conflict Resolution

When multiple devices would have the same designation, the script appends letter suffixes (A, B, C, etc.) similar to the wire numbering logic:
- `M1A1` - First motor at sheet 1, grid A1 (no suffix)
- `M1A1.A` - Second motor at sheet 1, grid A1
- `M1A1.B` - Third motor at sheet 1, grid A1

### Terminal Device Handling

Terminal devices are identified using the official E3 API methods `e3Device.IsTerminal()` and `e3Device.IsTerminalBlock()`, with fallback to letter code detection (T, TB, X, XT, TERM) if the API methods fail. Terminal devices receive special treatment:

1. **Complete skip** - Terminal devices are completely ignored by this script
2. **No designation changes** - Terminal devices keep their existing designations unchanged
3. **No processing** - Terminal devices are not processed for any designation-related operations

Terminal devices are identified and skipped entirely, leaving them untouched for other specialized scripts to handle.

## Files

- `apps/set_device_designations.py` - Main device designation automation script
- `apps/test_device_designation.py` - Test script to verify E3 connection and device operations
- `README_device_designation.md` - This documentation file

## Prerequisites

1. **E3.series** must be running with a project open
2. **Python** with `win32com.client` module (pywin32 package)
3. **Device symbols** must be placed on sheets (unplaced symbols are ignored)

## Device Letter Code Detection

The script attempts to determine device letter codes using the following methods (in order):

1. **Attribute-based detection** - Checks for common attributes:
   - "Device Letter Code"
   - "Letter Code"
   - "Type Code" 
   - "Device Type"

2. **Name-based detection** - Analyzes device names for common patterns:
   - Motor/MTR → "M"
   - Contactor/CONT → "K"
   - Transformer/XFMR → "T"
   - Relay → "K"
   - Switch → "S"
   - Fuse → "F"
   - Breaker/CB → "CB"

3. **Fallback** - Uses first letter of device name or "X" if unavailable

## Position Detection

The script finds the topmost leftmost symbol for each device by:

1. Getting all placed symbols for the device using `GetSymbolIds(1)`
2. For each symbol, getting position using `GetSchemaLocation()`
3. Comparing Y coordinates (topmost) then X coordinates (leftmost)
4. Extracting sheet assignment and grid position from the best symbol

## Usage

### Step 1: Test Connection

First, run the test script to verify E3 connection and examine device data:

```bash
python apps/test_device_designation.py
```

This will:
- Test connection to E3
- Display information about the first few devices
- Show available attributes and symbol data
- Help identify any configuration issues

### Step 2: Run Device Designation Assignment

Once the test is successful, run the main script:

```bash
python apps/set_device_designations.py
```

The script will:
1. Connect to the open E3 project
2. Get all devices in the project
3. For each device:
   - Determine the device letter code
   - Find the topmost leftmost symbol position
   - Generate base designation
4. Resolve designation conflicts with letter suffixes
5. Update device designation attributes

## Configuration

You may need to adjust the script for your specific E3 setup:

### Device Letter Code Attributes

In `get_device_letter_code()`, modify the `possible_attributes` list:

```python
possible_attributes = [
    "Device Letter Code",    # Add your attribute names here
    "Letter Code", 
    "Type Code",
    "Device Type"
]
```

### Device Designation Attributes

In `update_device_designation()`, modify the `possible_attributes` list:

```python
possible_attributes = [
    "Designation",          # Add your attribute names here
    "Device Designation", 
    "Name",
    "Reference"
]
```

### Device Type Patterns

In `get_device_letter_code()`, modify the name-based detection patterns:

```python
if "MOTOR" in name_upper or "MTR" in name_upper:
    return "M"
elif "CONTACTOR" in name_upper or "CONT" in name_upper:
    return "K"
# Add more patterns as needed
```

## Logging

The script creates detailed logs in `device_designation.log` including:
- Device processing details
- Position calculations
- Designation assignments
- Error messages and warnings

## API Methods Used

The script uses the following E3 DTM API methods:

### Device and Symbol Methods
- `e3Job.GetAllDeviceIds()` - Get all device IDs
- `e3Device.SetId()` - Set active device
- `e3Device.GetName()` - Get device name
- `e3Device.GetAssignment()` - Get device assignment
- `e3Device.GetSymbolIds()` - Get device symbol IDs

- `e3Device.GetComponentAttributeValue()` - Get device attributes
- `e3Device.SetName()` - Set device designation
- `e3Symbol.SetId()` - Set active symbol
- `e3Symbol.GetSchemaLocation()` - Get symbol position
- `e3Sheet.SetId()` - Set active sheet
- `e3Sheet.GetAssignment()` - Get sheet page number

### Terminal Device Detection Methods
- `e3Device.IsTerminal()` - Check if device is a terminal (returns 1 if true, 0 if false/error)
- `e3Device.IsTerminalBlock()` - Check if device is a terminal block/strip (returns 1 if true, 0 if false/error)

## Troubleshooting

### Common Issues

1. **No devices found**
   - Verify E3 project is open
   - Check that devices exist in the project

2. **Symbol position errors**
   - Ensure device symbols are placed on sheets
   - Check that symbols have valid grid positions

3. **Attribute errors**
   - Verify attribute names match your E3 setup
   - Use the test script to identify available attributes

4. **Designation not updating**
   - Check designation attribute names
   - Verify write permissions in E3 project

5. **Terminal device identification issues**
   - Check if E3 API methods `IsTerminal()` and `IsTerminalBlock()` are working correctly
   - Verify device letter code attribute values for fallback detection
   - Use logging to verify which devices are identified as terminals and skipped
   - Check debug logs for API method results vs fallback method results

### Debug Steps

1. Run the test script first to identify issues
2. Check the log file for detailed error messages
3. Verify E3 API object creation and method calls
4. Test with a small subset of devices first

## Integration

This script can be integrated with other E3 automation tools and follows the same patterns as the wire numbering script for consistency.
