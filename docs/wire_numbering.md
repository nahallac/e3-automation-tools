# E3 Wire Number Assignment Tool

This tool automatically assigns wire numbers to all net segments in an E3.series project based on page number and grid/ladder position.

## Overview

The wire numbering tool reads through the current open E3 project and for every connection, calculates wire numbers and sets the "Wire number" attribute on the corresponding net segments based on:

- **Page number**: Retrieved from the sheet assignment
- **Grid/ladder position**: Retrieved from the pin's schema location

The wire number format is: `{page_number}{grid_position}`

### Key Features

- **Lowest wire number selection**: For connections with multiple ends, the end that results in the lowest wire number is used
- **Signal-based unique numbering**: Each signal gets a unique wire number, with ALL segments and connections of the same signal receiving the same wire number. Letter suffixes (.A, .B, .C, etc.) are assigned left-to-right based on X coordinate when multiple signals share the same base wire number
- **FixWireName attribute support**: Connections whose nets have the "FixWireName" attribute set are automatically skipped and their wire numbers remain unchanged
- **Comprehensive logging**: Detailed logging of all operations and errors

## Files

- `set_wire_numbers.py` - Main script that performs the wire number assignment
- `test_fix_wire_name.py` - Test script to verify FixWireName attribute functionality
- `README_wire_numbering.md` - This documentation file

## Prerequisites

1. **E3.series** must be installed and running
2. **Python** with the following packages:
   - `e3series` (modern E3.series Python wrapper)
   - `pywin32` (legacy COM interface, kept for compatibility)
   - `logging` (built-in)
   - `sys` (built-in)
   - `collections` (built-in)
   - `re` (built-in)

3. **E3 Project** must be open in E3.series

## Installation

1. Ensure E3.series is installed and running
2. Install Python dependencies:
   ```bash
   pip install e3series pywin32
   ```
3. Copy the script files to your desired location

## Usage


### Step : Run Wire Number Assignment

Once the test is successful, run the main script:

```bash
python set_wire_numbers.py
```

The script will:
1. Connect to the open E3 project
2. Analyze all connections and their pin locations
3. Group all connections by signal name
4. Calculate base wire numbers for each signal based on lowest page and grid position
5. Get all net segments and X/Y coordinates for each signal
6. Group signals by base wire number
7. Sort signals within each group by X coordinate (left-to-right), then Y coordinate (top-to-bottom)
8. Assign unique wire numbers with letter suffixes based on position order
9. Update ALL net segments belonging to each signal with the same wire number

## Wire Number Calculation

### Format
```
{page_number}{grid_position}
```

### Examples
- Page 1, Grid A5: `1A5`
- Page 10, Grid B12: `10B12`
- Page 2, Column C, Row 7: `2C7`
- Multiple signals at same position (left-to-right order):
  - Leftmost signal at 1A5: `1A5`
  - Next signal to the right at 1A5: `1A5.A`
  - Third signal to the right at 1A5: `1A5.B`

### Grid Position Extraction

The script attempts to extract grid position in this order:
1. From grid description (format: `/sheet.grid`)
2. From column and row values combined
3. From column value only
4. From row value only
5. Fallback to "UNKNOWN"

### Signal-Based Wire Number Assignment

For each signal in the project, the script:
1. Groups all connections by signal name
2. For each signal, calculates base wire numbers for all pin locations
3. Sorts them lexicographically and selects the lowest value as the base
4. Groups signals by their base wire number
5. Within each group, sorts signals by X coordinate (left-to-right), then Y coordinate (top-to-bottom)
6. Assigns wire numbers based on position order:
   - First signal (leftmost): gets the base wire number
   - Second signal: gets base wire number + `.A`
   - Third signal: gets base wire number + `.B`
   - And so on through the alphabet
7. **ALL segments and connections belonging to the same signal receive the same wire number**

### Letter Suffix Logic

When multiple signals share the same base wire number, letter suffixes are assigned based on physical position:
- Leftmost signal: no suffix (e.g., `1A5`)
- Next signal to the right: `.A` (e.g., `1A5.A`)
- Third signal from left: `.B` (e.g., `1A5.B`)
- Fourth signal from left: `.C` (e.g., `1A5.C`)
- And so on through the alphabet

If signals have the same X coordinate, they are sorted by Y coordinate (top-to-bottom).

This ensures that every signal in the project has a completely unique wire number, and the letter suffixes follow a logical left-to-right progression across the schematic page.


## Logging

The script creates a detailed log file: `wire_numbering.log`

Log levels:
- **INFO**: General progress and successful operations
- **WARNING**: Non-critical issues (e.g., pins without schema location)
- **ERROR**: Critical errors that prevent processing
- **DEBUG**: Detailed debugging information (when enabled)

## Error Handling

Common issues and solutions:

### "Failed to connect to E3"
- Ensure E3.series is running
- Verify a project is open in E3
- Check that no other applications are blocking COM access

### "No connections found in project"
- Verify the project contains electrical connections
- Check that the project is properly loaded

### "Could not set wire number for connection"
- The project may use different attribute names for wire numbers
- Check the E3 project's attribute definitions
- Manually verify attribute names in E3

### "Pin has no schema location"
- Some pins may not be placed on schematic sheets
- This is normal for certain types of connections
- These connections will be skipped

## FixWireName Attribute

The script supports the "FixWireName" net attribute to exclude specific connections from automatic wire number assignment. This is useful when you have manually assigned wire numbers that should not be changed.

### How to Use

1. **In E3.series**: Select the net you want to protect
2. **Set the attribute**: Add a net attribute called "FixWireName" with any non-empty value (e.g., "1", "true", "yes")
3. **Run the script**: The script will automatically detect and skip all connections on this net

### Attribute Values

The script will skip connections if the "FixWireName" attribute is set to any value except:
- Empty string (`""`)
- `"0"`
- `"false"` (case-insensitive)
- `"no"` (case-insensitive)

### Example

If you have a net with wire number "POWER_24V" that should never be changed:
1. Select the net in E3
2. Add attribute: `FixWireName = "1"`
3. The script will skip all connections on this net and preserve "POWER_24V"

### Testing FixWireName Functionality

To test which connections in your project have the FixWireName attribute set:

```bash
python test_fix_wire_name.py
```

This test script will:
- Connect to your open E3 project
- Check the first 10 connections for the FixWireName attribute
- Report which connections would be skipped during wire numbering
- Generate a detailed log file `test_fix_wire_name.log`

## Customization

### Modifying Wire Number Format

To change the wire number format, edit the `calculate_wire_number` method in `set_wire_numbers.py`:

```python
def calculate_wire_number(self, page_number, grid_position):
    # Current format: page_number + grid_position
    wire_number = f"{page_num}{grid_position}"

    # Example alternative formats:
    # wire_number = f"{page_num}-{grid_position}"  # Dash separator
    # wire_number = f"W{page_num}{grid_position}"  # With prefix
    # wire_number = f"{page_num}({grid_position})"  # With parentheses

    return wire_number
```

### Adding Custom Attribute Names

To add custom wire number attribute names, edit the attribute list in the `process_connections` method:

```python
for attr_name in ["Wire number", "WIRE_NUMBER", "WireNumber", "Wire_Number", "YOUR_CUSTOM_NAME"]:
```

## Technical Details

### DTM API Methods Used

- `e3Job.GetAllConnectionIds()` - Get all connection IDs
- `e3Connection.SetId()` - Set active connection
- `e3Connection.GetSignalName()` - Get signal name
- `e3Connection.GetPinIds()` - Get connection pin IDs
- `e3Connection.GetNetSegmentIds()` - Get net segment IDs for connection
- `e3NetSegment.SetId()` - Set active net segment
- `e3NetSegment.SetAttributeValue()` - Set wire number attribute on net segment
- `e3Pin.SetId()` - Set active pin
- `e3Pin.GetSchemaLocation()` - Get pin location information
- `e3Sheet.SetId()` - Set active sheet
- `e3Sheet.GetAssignment()` - Get sheet page number

### COM Object Management

The script properly manages COM objects by:
- Creating objects through the E3 application
- Setting objects to None for cleanup
- Using try/except blocks for error handling

## Version Information

- **Script Version**: 1.0
- **E3.series Compatibility**: v2023-24.30 (and likely earlier versions)
- **Python Version**: 3.6+

## Support

For issues or questions:
1. Check the log file for detailed error information
2. Verify E3 project structure and attribute definitions
4. Ensure all prerequisites are met

## License

This script is provided as-is for educational and automation purposes. Use at your own risk and always backup your E3 projects before running automated scripts.
