# E3 Terminal Pin Name Automation

This script automatically sets terminal pin names to match the wire numbers of their connected net segments in E3.series projects.

## Overview

The terminal pin name automation script processes all terminal devices in an E3.series project and updates each pin's name to match the "Wire number" attribute of the connected net segment. This ensures consistency between wire numbering and terminal pin naming.

### Key Features

- **Automatic Terminal Detection**: Uses E3 API methods `IsTerminal()` and `IsTerminalBlock()` to identify terminal devices
- **Wire Number Retrieval**: Gets wire numbers from connected net segments using the "Wire number" attribute
- **Pin Name Updates**: Sets terminal pin names to match their corresponding wire numbers
- **Comprehensive Logging**: Detailed logging of all operations and errors
- **Modern API**: Uses the e3series PyPI library for improved reliability
- **Error Handling**: Robust error handling for various edge cases

## Files

- `set_terminal_pin_names.py` - Main script that performs the terminal pin name automation
- `README_terminal_pin_names.md` - This documentation file

## Prerequisites

1. **E3.series** must be installed and running
2. **Python** with the following packages:
   - `e3series` (modern E3.series Python wrapper)
   - `logging` (built-in)
   - `sys` (built-in)
   - `collections` (built-in)

3. **E3 Project** must be open in E3.series with:
   - Terminal devices present in the project
   - Wire numbers assigned to net segments (use the wire numbering script first if needed)

## Installation

1. Ensure E3.series is installed and running
2. Install Python dependencies:
   ```bash
   pip install e3series
   ```
3. Copy the script files to your desired location

## Usage

### Basic Usage

1. **Open E3.series** and load your project
2. **Ensure wire numbers are assigned** to net segments (run the wire numbering script if needed)
3. **Run the script**:
   ```bash
   python set_terminal_pin_names.py
   ```
4. **Check the results** in the log file and E3.series project

### What the Script Does

1. **Connects** to the active E3.series application
2. **Identifies** all terminal devices using E3 API methods
3. **Gets all pins** from each terminal device
4. **Retrieves wire numbers** from connected net segments
5. **Updates pin names** to match wire numbers
6. **Logs all operations** for review and debugging

## Script Logic

### Terminal Device Detection

The script uses the official E3 API methods to identify terminal devices:
- `e3Device.IsTerminal()` - Checks if device is a terminal
- `e3Device.IsTerminalBlock()` - Checks if device is a terminal block

### Pin Processing

For each terminal device:
1. Get all pin IDs using `e3Device.GetPinIds()`
2. For each pin, get connected net segments using `e3Pin.GetNetSegmentIds()`
3. Retrieve wire number from net segment using `e3NetSegment.GetAttributeValue("Wire number")`
4. Set pin name using `e3Pin.SetName(wire_number)`

### Error Handling

The script handles various scenarios:
- Pins with no connected net segments
- Net segments with no wire numbers
- Pins that already have the correct name
- API errors and connection issues

## Output and Logging

The script creates a detailed log file `terminal_pin_names.log` containing:
- Connection status and initialization
- Terminal device detection results
- Pin processing details
- Success/failure status for each operation
- Summary statistics

### Log Levels

- **INFO**: General progress and summary information
- **DEBUG**: Detailed operation information
- **WARNING**: Non-critical issues (e.g., pins with no wire numbers)
- **ERROR**: Critical errors that prevent operation

## Example Output

```
2025-01-11 10:30:15 - INFO - Starting E3 Terminal Pin Name Automation
2025-01-11 10:30:15 - INFO - Successfully connected to E3 application using e3series library
2025-01-11 10:30:16 - INFO - Found 5 terminal devices out of 25 total devices
2025-01-11 10:30:16 - INFO - Processing terminal device TB1 (12345) with 12 pins
2025-01-11 10:30:16 - INFO - Updated pin 67890: '1' -> '101A'
2025-01-11 10:30:16 - INFO - Updated pin 67891: '2' -> '101B'
2025-01-11 10:30:17 - INFO - Updated 10/12 pins for device TB1
2025-01-11 10:30:18 - INFO - Terminal pin processing complete: 48/60 pins updated
2025-01-11 10:30:18 - INFO - E3 Terminal Pin Name Automation completed successfully
```

## Integration with Other Scripts

This script works best when used in conjunction with other E3 automation tools:

1. **Wire Numbering Script**: Run first to assign wire numbers to net segments
2. **Device Designation Script**: Can be run before or after this script
3. **Publishing Tools**: Run after all automation to finalize the project

## API Methods Used

### E3 Device Methods
- `e3Job.GetAllDeviceIds()` - Get all device IDs in project
- `e3Device.SetId()` - Set active device
- `e3Device.IsTerminal()` - Check if device is a terminal
- `e3Device.IsTerminalBlock()` - Check if device is a terminal block
- `e3Device.GetName()` - Get device name
- `e3Device.GetPinIds()` - Get device pin IDs

### E3 Pin Methods
- `e3Pin.SetId()` - Set active pin
- `e3Pin.GetName()` - Get pin name
- `e3Pin.SetName()` - Set pin name
- `e3Pin.GetNetSegmentIds()` - Get connected net segment IDs

### E3 Net Segment Methods
- `e3NetSegment.SetId()` - Set active net segment
- `e3NetSegment.GetAttributeValue()` - Get wire number attribute

## Troubleshooting

### Common Issues

1. **No terminal devices found**
   - Verify terminal devices exist in the project
   - Check that devices are properly identified as terminals in E3

2. **Pins not updated**
   - Ensure wire numbers are assigned to net segments
   - Check that pins are connected to net segments
   - Verify pin names are modifiable (some pins may be read-only)

3. **Connection errors**
   - Ensure E3.series is running and a project is open
   - Check that the e3series Python package is installed
   - Verify no other scripts are accessing E3 simultaneously

4. **Permission errors**
   - Check that the project is not read-only
   - Verify user permissions in E3.series
   - Ensure the project is not locked by another user

### Debug Steps

1. Run the script and check the log file for detailed error messages
2. Verify E3 API object creation and method calls
3. Test with a small subset of terminal devices first
4. Check wire number assignments manually in E3.series

## Version Information

- **Created**: January 2025
- **E3.series Compatibility**: v2009-8.50 and later
- **Python Requirements**: Python 3.6+
- **Dependencies**: e3series PyPI package

## Support

For issues with this script:
1. Check the log file for detailed error information
2. Verify all prerequisites are met
3. Test with the wire numbering script first
4. Review the E3.series project for data consistency

---

**Note**: This script is designed to work with projects that already have wire numbers assigned. If wire numbers are not present, run the wire numbering automation script first.
