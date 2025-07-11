# E3.series Automation Tools

A collection of automation scripts for E3.series electrical design software, focusing on wire numbering and device designation management.

## Features

### Current Tools
- **Wire Number Assignment** (`set_wire_numbers.py`) - Automatically assigns wire numbers based on page and grid position with intelligent conflict resolution

### Planned Tools
- **Device Designation Assignment** - Automatically assign device designations based on configurable rules
- **GUI Interface** - User-friendly interface for running and configuring all automation tools

## Quick Start

### Prerequisites
- E3.series software installed and running
- Python 3.7+ with the following packages:
  ```bash
  pip install e3series pywin32
  ```
- An open E3 project

### Wire Number Assignment

1. Open your E3 project
2. Run the wire numbering script:
   ```bash
   python scripts/set_wire_numbers.py
   ```
3. Check the generated log file `wire_numbering.log` for details

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
