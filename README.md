# E3.series Automation Tools

A comprehensive collection of automation scripts and GUI tools for E3.series electrical design software, focusing on North American electrical standards compliance.

## Features

### Current Tools
- **Wire Number Assignment** - Automatically assigns wire numbers based on page and grid position with intelligent conflict resolution
- **Device Designation Assignment** - Automatically assigns device designations based on sheet and grid position
- **Terminal Pin Name Assignment** - Sets terminal pin names to match connected wire numbers
- **NA Standards GUI** - User-friendly interface for running all automation tools with real-time feedback

### Available Interfaces
- **Individual Scripts** - Run each automation independently from command line
- **Library Modules** - Import and use automation functions in custom workflows
- **GUI Application** - Modern interface with batch processing and live logging

## Quick Start

### Prerequisites
- E3.series software installed and running
- Python 3.7+ with required packages (see Installation section)
- An open E3 project

### Using the GUI (Recommended)

1. Open your E3 project
2. Run the NA Standards GUI:
   ```bash
   python gui/e3_NA_Standards.py
   ```
3. Use individual buttons or "Run All" for complete automation
4. Monitor progress in the real-time log area

### Using Individual Scripts

1. Wire numbering:
   ```bash
   python scripts/set_wire_numbers.py
   ```
2. Device designations:
   ```bash
   python lib/e3_device_designation.py
   ```
3. Terminal pin names:
   ```bash
   python lib/e3_terminal_pin_names.py
   ```

## How It Works

### Wire Numbering Logic
The wire numbering script:
1. Analyzes all connections in the open E3 project
2. Groups connections by signal name
3. Calculates wire numbers using format: `{page_number}{grid_position}`
4. Assigns the lowest possible wire number to each signal
5. Handles conflicts by adding letter suffixes (.A, .B, .C, etc.) based on left-to-right position

### Example
- Signal at page 1, grid A5 → Wire number: `1A5`
- Multiple signals at same position → `1A5`, `1A5.A`, `1A5.B` (sorted left to right)

## Installation

1. Download or clone this repository
2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure E3.series is running with a project open
4. Run the desired script from the `scripts/` directory

## Documentation

- [Wire Numbering Guide](docs/wire_numbering.md) - Detailed documentation for the wire numbering tool
- [Device Designation Guide](docs/device_designation.md) - Device designation automation documentation
- [Terminal Pin Names Guide](docs/terminal_pin_names.md) - Terminal pin naming automation documentation
- [NA Standards GUI Guide](docs/na_standards_gui.md) - Complete GUI application documentation
- [Installation Guide](docs/installation.md) - Step-by-step installation instructions

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

## License

This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.

You are free to share and adapt this work, but you must:
- Give appropriate credit to the original author
- Indicate if changes were made
- Distribute any derivative works under the same license

See the LICENSE file for full details.

## Author

Jonathan Callahan
