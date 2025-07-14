# E3 NA Standards GUI

The E3 NA Standards GUI provides a user-friendly interface for running E3.series automation scripts that comply with North American electrical standards.

## Features

- **Individual Script Execution**: Run each automation script separately
- **Batch Processing**: "Run All" button executes all scripts in the correct sequence
- **Real-time Feedback**: Live logging and status updates
- **Error Handling**: Comprehensive error reporting and recovery

## Available Operations

### 1. Set Wire Numbers
Automatically assigns wire numbers based on page number and grid position. Wire numbers follow the format `{page_number}{grid_position}` with letter suffixes for multiple signals at the same position.

### 2. Set Device Designations  
Updates device designations based on the sheet and grid position of the device's first symbol. Format: `{device_letter_code}{sheet}{grid}` with conflict resolution.

### 3. Set Terminal Pin Names
Sets terminal pin names to match the wire numbers of their connected net segments.

### 4. Run All (Recommended)
Executes all three operations in the optimal sequence:
1. Wire Numbers → 2. Device Designations → 3. Terminal Pin Names

## Requirements

- E3.series software installed and running
- Python 3.7 or higher
- Required packages (see requirements.txt):
  - `e3series` - E3.series Python API
  - `customtkinter` - Modern GUI framework
  - `pywin32` - Windows COM interface

## Installation

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Ensure E3.series is running with a project open

## Usage

1. Run the GUI application:
   ```bash
   python gui/e3_NA_Standards.py
   ```

2. The GUI will open with four main buttons:
   - **Set Wire Numbers**: Run wire numbering automation
   - **Set Device Designations**: Run device designation automation  
   - **Set Terminal Pin Names**: Run terminal pin name automation
   - **Run All**: Execute all three in sequence (recommended)

3. Click any button to start the operation. Progress will be shown in:
   - Status bar at the bottom
   - Real-time log area in the center
   - Button states (disabled during operation)

4. Use the "Clear Log" button to reset the log area

## Operation Sequence

When using "Run All", the scripts execute in this order:

1. **Wire Numbers First**: Establishes the wire numbering scheme
2. **Device Designations Second**: Uses wire positions for device naming
3. **Terminal Pin Names Last**: Matches pin names to established wire numbers

This sequence ensures consistency and prevents conflicts between the different automation systems.

## Error Handling

- **Connection Errors**: Verify E3.series is running with a project open
- **COM Errors**: The application automatically initializes COM interfaces
- **Operation Failures**: Check the log area for detailed error messages
- **Partial Failures**: "Run All" stops on first failure to prevent inconsistent states

## Logging

All operations are logged with timestamps showing:
- Connection status
- Processing progress  
- Success/failure notifications
- Detailed error information when issues occur

## Troubleshooting

**GUI won't start**: 
- Verify Python and required packages are installed
- Check that customtkinter is properly installed

**Can't connect to E3**:
- Ensure E3.series is running
- Verify a project is open in E3.series
- Check Windows COM permissions

**Operations fail**:
- Review the log area for specific error messages
- Ensure the E3 project is not read-only
- Verify sufficient permissions to modify the project

## Integration

This GUI integrates with the E3 automation library modules:
- `lib/e3_wire_numbering.py`
- `lib/e3_device_designation.py` 
- `lib/e3_terminal_pin_names.py`

The library modules can also be used independently for custom automation workflows.
