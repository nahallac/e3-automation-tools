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
- **Automatic attribute detection**: Tries multiple common wire number attribute names
- **Comprehensive logging**: Detailed logging of all operations and errors

## Usage

### Basic Usage

1. Open your E3 project in E3.series
2. Run the wire numbering script:
   ```bash
   python scripts/set_wire_numbers.py
   ```
3. Check the generated log file `wire_numbering.log` for details

### What the Script Does

1. **Connects to E3**: Establishes COM connection to the running E3 application
2. **Analyzes Connections**: Reads all connections in the project
3. **Groups by Signal**: Groups connections that belong to the same signal
4. **Calculates Wire Numbers**: Determines the optimal wire number for each signal
5. **Handles Conflicts**: Assigns letter suffixes when multiple signals share the same base wire number
6. **Updates Net Segments**: Sets the "Wire number" attribute on all relevant net segments

## Wire Number Format

### Basic Format
- **Page 1, Grid A5** → Wire number: `1A5`
- **Page 10, Grid B12** → Wire number: `10B12`

### Conflict Resolution
When multiple signals share the same grid position, letter suffixes are added based on left-to-right position:
- **First signal** → `1A5`
- **Second signal** → `1A5.A`
- **Third signal** → `1A5.B`

## Logging

The script creates a detailed log file: `wire_numbering.log`

Log levels:
- **INFO**: General progress and successful operations
- **WARNING**: Non-critical issues (e.g., pins without schema location)
- **ERROR**: Critical errors that prevent processing
- **DEBUG**: Detailed debugging information (when enabled)

## Error Handling

### Common Issues and Solutions

#### "Failed to connect to E3"
- Ensure E3.series is running
- Verify a project is open in E3
- Check that no other applications are blocking COM access

#### "No connections found in project"
- Verify the project has electrical connections
- Check that the project is properly loaded in E3

#### "Could not calculate wire number for signal"
- Check that pins have valid schema locations
- Verify grid/ladder setup in your E3 project

## Technical Details

### E3 API Objects Used
- **Connection Object**: For reading connection information
- **Pin Object**: For getting pin locations and grid positions
- **Sheet Object**: For retrieving page numbers
- **Signal Object**: For grouping connections by signal
- **NetSegment Object**: For setting wire number attributes

### Attribute Names
The script attempts to set wire numbers using these attribute names (in order):
1. `Wire number`
2. `WIRE_NUMBER`
3. `WireNumber`
4. `Wire_Number`

The first successful attribute name is used for all connections.

## Troubleshooting

### Performance Issues
- Large projects may take several minutes to process
- Monitor the log file for progress updates
- Consider running on smaller sections if needed

### Unexpected Wire Numbers
- Check the log file for detailed calculation information
- Verify your E3 project's grid/ladder setup
- Ensure page numbers are correctly assigned

### COM Interface Issues
- Restart E3.series if COM errors occur
- Close other applications that might use E3's COM interface
- Run the script as Administrator if permission issues occur
