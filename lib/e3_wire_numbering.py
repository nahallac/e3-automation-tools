#!/usr/bin/env python3
"""
E3 Wire Number Assignment Library

This module provides functionality to read through E3 projects and for every connection,
calculate wire numbers based on page number and grid/ladder position, then set the 
"Wire number" attribute on the corresponding net segments.

The wire number format is: {page_number}{grid_position}
The end that results in the lowest wire number is used, with proper numeric
sorting that considers both page numbers and grid positions numerically.
Each signal gets a unique wire number, and ALL segments and connections
belonging to the same signal receive the same wire number. When multiple
signals share the same base wire number (same grid position), letter
suffixes (.A, .B, .C, etc.) are assigned based on left-to-right position.

Connections with the "FixWireName" net attribute set will be skipped and
their wire numbers will not be modified.

Author: Jonathan Callahan
Date: 2025-07-01
Updated: 2025-07-08 - Migrated to use e3series PyPI package instead of win32com.client
Updated: 2025-07-11 - Added support for FixWireName attribute to skip connections
Updated: 2025-07-11 - Fixed wire number sorting to properly consider both page and grid numerically
Updated: 2025-07-11 - Converted to library module for GUI integration
"""

import e3series
import logging
import sys
import pythoncom

class WireNumberAssigner:
    def __init__(self, logger=None, e3_pid=None):
        self.app = None
        self.job = None
        self.connection = None
        self.pin = None
        self.sheet = None
        self.signal = None
        self.net = None
        self.net_segment = None
        self.logger = logger or logging.getLogger(__name__)
        self.e3_pid = e3_pid
        
    def connect_to_e3(self):
        """Connect to E3 application using connection manager"""
        try:
            # Get E3 PID if not already provided
            if self.e3_pid is None:
                from .e3_connection_manager import get_e3_connection_pid
                self.e3_pid = get_e3_connection_pid(self.logger)
                if self.e3_pid is None:
                    return False

            # Connect using the PID
            from .e3_connection_manager import connect_to_e3_with_pid
            success, objects = connect_to_e3_with_pid(self.e3_pid, self.logger)
            if success:
                self.app = objects['app']
                self.job = objects['job']
                self.connection = objects['connection']
                self.pin = objects['pin']
                self.sheet = objects['sheet']
                self.signal = objects['signal']
                self.net = objects['net']
                self.net_segment = objects['net_segment']
                return True
            else:
                return False
        except Exception as e:
            self.logger.error(f"Failed to connect to E3: {e}")
            return False
    
    def get_pin_location_info(self, pin_id):
        """Get location information for a pin"""
        try:
            self.pin.SetId(pin_id)

            # Get schema location - E3 API returns (sheet_id, x, y, grid_desc, column, row)
            try:
                result = self.pin.GetSchemaLocation()

                if isinstance(result, tuple) and len(result) >= 6:
                    sheet_id = result[0]
                    x_coord = result[1]  # x coordinate (used for left-to-right sorting)
                    y_coord = result[2]  # y coordinate (used for top-to-bottom sorting)
                    grid_desc = result[3]
                    column = result[4]
                    row = result[5]
                else:
                    self.logger.warning(f"Unexpected GetSchemaLocation result format for pin {pin_id}: {result}")
                    return None, None, None, None, None

            except Exception as e:
                self.logger.error(f"Error calling GetSchemaLocation for pin {pin_id}: {e}")
                return None, None, None, None, None

            if not sheet_id or sheet_id <= 0:
                self.logger.debug(f"Pin {pin_id} has no valid schema location (sheet_id: {sheet_id})")
                return None, None, None, None, None


            try:
                self.sheet.SetId(sheet_id)
                page_number = self.sheet.GetName()
            except Exception as e:
                self.logger.error(f"Error getting sheet assignment for sheet {sheet_id}: {e}")
                page_number = "UNKNOWN"

            # Extract grid position from grid_desc or use column/row
            grid_position = self.extract_grid_position(grid_desc, column, row)

            self.logger.debug(f"Pin {pin_id}: Sheet {sheet_id}, Page {page_number}, Grid {grid_position}, X={x_coord}, Y={y_coord}")

            return page_number, grid_position, sheet_id, x_coord, y_coord

        except Exception as e:
            self.logger.error(f"Error getting pin location for pin {pin_id}: {e}")
            return None, None, None, None, None
    
    def extract_grid_position(self, grid_desc, column, row):
        """Extract grid position from grid description or column/row"""
        try:
            # If we have grid_desc in format "/sheet.grid", extract the grid part
            if grid_desc and "." in grid_desc:
                grid_part = grid_desc.split(".")[-1]
                return grid_part

            # If we have column and row, combine them
            if column and row:
                return f"{column}{row}"

            # Fallback to just column or row if available
            if column:
                return column
            if row:
                return row

            return "UNKNOWN"

        except Exception as e:
            self.logger.error(f"Error extracting grid position: {e}")
            return "UNKNOWN"

    def wire_number_sort_key(self, wire_num):
        """Create a sort key that handles numeric page and grid comparison"""
        try:
            # Extract page and grid parts from wire number
            # Wire number format is: {page_number}{grid_position}
            # We need to separate numeric page from alphanumeric grid

            # Find where the page number ends and grid starts
            # Page should be numeric, grid can be alphanumeric
            page_part = ""
            grid_part = ""

            i = 0
            # Extract numeric page part
            while i < len(wire_num) and wire_num[i].isdigit():
                page_part += wire_num[i]
                i += 1

            # Rest is grid part
            grid_part = wire_num[i:]

            # Convert page to int for proper numeric sorting, default to 0 if empty
            page_num = int(page_part) if page_part else 0

            # For grid part, we need to handle complex patterns like:
            # "A", "A1", "A10", "B2", "C1", etc.
            # We'll break it down into: leading_letters + numbers + trailing_letters

            grid_leading_alpha = ""
            grid_num_part = ""
            grid_trailing_alpha = ""

            j = 0
            # Extract leading letters
            while j < len(grid_part) and grid_part[j].isalpha():
                grid_leading_alpha += grid_part[j]
                j += 1

            # Extract numbers
            while j < len(grid_part) and grid_part[j].isdigit():
                grid_num_part += grid_part[j]
                j += 1

            # Rest is trailing letters
            grid_trailing_alpha = grid_part[j:]

            grid_num = int(grid_num_part) if grid_num_part else 0

            return (page_num, grid_leading_alpha, grid_num, grid_trailing_alpha)

        except Exception as e:
            self.logger.warning(f"Error parsing wire number '{wire_num}' for sorting: {e}")
            # Fallback to lexicographic sorting for this wire number
            return (999999, 'ZZZ', 999999, wire_num)
    
    def calculate_wire_number(self, page_number, grid_position):
        """Calculate wire number from page and grid position"""
        try:
            # Handle empty or None page numbers
            if not page_number or page_number.strip() == "":
                page_num = "0"
            else:
                page_num = str(page_number).strip()
            
            # Format: page_number + grid_position
            wire_number = f"{page_num}{grid_position}"
            return wire_number
            
        except Exception as e:
            self.logger.error(f"Error calculating wire number: {e}")
            return "ERROR"
    
    def get_connection_wire_numbers_and_positions(self, connection_id):
        """Get potential wire numbers and positions for both ends of a connection"""
        try:
            self.connection.SetId(connection_id)

            # Get pin IDs for this connection
            pin_ids_result = self.connection.GetPinIds()
            if not pin_ids_result:
                self.logger.warning(f"Connection {connection_id} has no pins")
                return []

            wire_data = []

            # E3 API returns (count, tuple_of_ids) for pin IDs too
            actual_pin_ids = []
            if isinstance(pin_ids_result, tuple) and len(pin_ids_result) >= 2:
                _ = pin_ids_result[0]  # count (not used)
                pin_ids = pin_ids_result[1]

                if isinstance(pin_ids, tuple):
                    # Filter out None values and invalid pin IDs (like 0)
                    actual_pin_ids = [pid for pid in pin_ids if pid is not None and pid != 0]
                else:
                    if pin_ids is not None and pin_ids != 0:
                        actual_pin_ids = [pin_ids]
            else:
                self.logger.warning(f"Unexpected pin IDs format for connection {connection_id}: {type(pin_ids_result)}")
                return []

            self.logger.debug(f"Connection {connection_id} has {len(actual_pin_ids)} valid pins")

            for pin_id in actual_pin_ids:
                page_number, grid_position, _, x_coord, y_coord = self.get_pin_location_info(pin_id)
                if page_number is not None and grid_position is not None:
                    wire_number = self.calculate_wire_number(page_number, grid_position)
                    wire_data.append({
                        'wire_number': wire_number,
                        'x_coord': x_coord if x_coord is not None else 0,
                        'y_coord': y_coord if y_coord is not None else 0,
                        'pin_id': pin_id
                    })
                    self.logger.debug(f"Pin {pin_id}: Page {page_number}, Grid {grid_position} -> Wire {wire_number}, X={x_coord}, Y={y_coord}")

            return wire_data

        except Exception as e:
            self.logger.error(f"Error getting wire numbers for connection {connection_id}: {e}")
            return []

    def get_connection_wire_numbers(self, connection_id):
        """Get potential wire numbers for both ends of a connection (backward compatibility)"""
        wire_data = self.get_connection_wire_numbers_and_positions(connection_id)
        return [data['wire_number'] for data in wire_data]
    
    def get_lowest_wire_number(self, wire_numbers):
        """Get the lowest wire number from a list"""
        if not wire_numbers:
            return None

        # Sort wire numbers using the custom key that handles numeric page and grid comparison
        sorted_numbers = sorted(wire_numbers, key=self.wire_number_sort_key)
        return sorted_numbers[0]

    def get_net_segments_for_connection(self, connection_id):
        """Get all net segment IDs for a given connection"""
        try:
            self.connection.SetId(connection_id)
            net_segment_ids_result = self.connection.GetNetSegmentIds()

            if not net_segment_ids_result:
                return []

            actual_net_segments = []
            if isinstance(net_segment_ids_result, tuple) and len(net_segment_ids_result) >= 2:
                _ = net_segment_ids_result[0]  # count (not used)
                net_segment_ids = net_segment_ids_result[1]

                if isinstance(net_segment_ids, tuple):
                    actual_net_segments = [nsid for nsid in net_segment_ids if nsid is not None and nsid != 0]
                else:
                    if net_segment_ids is not None and net_segment_ids != 0:
                        actual_net_segments = [net_segment_ids]
            else:
                self.logger.warning(f"Unexpected net segment IDs format for connection {connection_id}: {type(net_segment_ids_result)}")

            return actual_net_segments

        except Exception as e:
            self.logger.error(f"Error getting net segments for connection {connection_id}: {e}")
            return []

    def has_fix_wire_name_attribute(self, connection_id):
        """Check if the net for this connection has the FixWireName attribute set"""
        try:
            self.connection.SetId(connection_id)
            net_id = self.connection.GetNetId()

            # Check if we got a valid net ID
            if net_id <= 0:
                self.logger.debug(f"Connection {connection_id} has no valid net ID ({net_id}) - processing normally")
                return False

            # Set the net object to this net and check for FixWireName attribute
            self.net.SetId(net_id)
            fix_wire_name = self.net.GetAttributeValue("FixWireName")

            # Check if the attribute exists and has a truthy value
            if fix_wire_name and str(fix_wire_name).strip().lower() not in ['', '0', 'false', 'no']:
                self.logger.debug(f"Connection {connection_id} has FixWireName attribute set to '{fix_wire_name}' on net {net_id} - skipping")
                return True

            return False

        except Exception as e:
            self.logger.error(f"Error checking FixWireName attribute for connection {connection_id}: {e}")
            return False

    def process_connections(self):
        """Process all connections in the project"""
        try:
            # Get all connection IDs
            connection_ids_result = self.job.GetAllConnectionIds()
            if not connection_ids_result:
                self.logger.warning("No connections found in project")
                return

            # E3 API returns (count, tuple_of_ids)
            actual_connections = []
            if isinstance(connection_ids_result, tuple) and len(connection_ids_result) >= 2:
                count = connection_ids_result[0]
                connection_ids = connection_ids_result[1]

                self.logger.info(f"E3 reports {count} connections")

                if isinstance(connection_ids, tuple):
                    # Filter out None values
                    actual_connections = [cid for cid in connection_ids if cid is not None]
                else:
                    actual_connections = [connection_ids] if connection_ids is not None else []
            else:
                self.logger.warning(f"Unexpected connection IDs format: {type(connection_ids_result)}")
                return

            self.logger.info(f"Found {len(actual_connections)} valid connections to process")

            # Group connections by signal name first
            from collections import defaultdict
            signal_connections = defaultdict(list)

            # First pass: group all connections by signal name
            for conn_id in actual_connections:
                if conn_id is None:
                    continue

                try:
                    # Check if this connection has the FixWireName attribute set
                    if self.has_fix_wire_name_attribute(conn_id):
                        self.logger.info(f"Skipping connection {conn_id} - has FixWireName attribute set")
                        continue

                    self.connection.SetId(conn_id)
                    signal_name = self.connection.GetSignalName()

                    if signal_name:  # Only process connections with valid signal names
                        signal_connections[signal_name].append(conn_id)
                        self.logger.debug(f"Connection {conn_id} belongs to signal '{signal_name}'")

                except Exception as e:
                    self.logger.error(f"Error getting signal name for connection {conn_id}: {e}")

            # Second pass: calculate base wire number for each signal
            signal_data = []

            for signal_name, connection_ids in signal_connections.items():
                try:
                    # Collect all wire numbers and positions for this signal
                    all_wire_data = []
                    all_net_segments = []

                    for conn_id in connection_ids:
                        # Get wire numbers and positions for this connection
                        wire_data = self.get_connection_wire_numbers_and_positions(conn_id)
                        all_wire_data.extend(wire_data)

                        # Get all net segments for this connection
                        net_segment_ids = self.get_net_segments_for_connection(conn_id)
                        all_net_segments.extend(net_segment_ids)

                    if all_wire_data:
                        # Find the wire data with the lowest wire number for this signal
                        # Use the same sorting logic as get_lowest_wire_number
                        all_wire_data.sort(key=lambda x: self.wire_number_sort_key(x['wire_number']))
                        lowest_wire_data = all_wire_data[0]

                        # Store signal data for later processing
                        signal_data.append({
                            'signal_name': signal_name,
                            'base_wire_number': lowest_wire_data['wire_number'],
                            'x_coord': lowest_wire_data['x_coord'],
                            'y_coord': lowest_wire_data['y_coord'],
                            'net_segment_ids': list(set(all_net_segments)),  # Remove duplicates
                            'connection_ids': connection_ids
                        })

                        self.logger.debug(f"Signal '{signal_name}' -> Base Wire: {lowest_wire_data['wire_number']}, X={lowest_wire_data['x_coord']}, Connections: {len(connection_ids)}, Net Segments: {len(set(all_net_segments))}")
                    else:
                        self.logger.warning(f"Could not calculate wire number for signal '{signal_name}'")

                except Exception as e:
                    self.logger.error(f"Error processing signal '{signal_name}': {e}")

            # Group signals by base wire number
            wire_number_groups = defaultdict(list)
            for signal_info in signal_data:
                wire_number_groups[signal_info['base_wire_number']].append(signal_info)

            # Sort each group by X coordinate (left to right) then Y coordinate (top to bottom)
            for base_wire_number, signals in wire_number_groups.items():
                signals.sort(key=lambda x: (x['x_coord'], x['y_coord']))

            # Third pass: assign unique wire numbers to each signal
            updated_count = 0
            used_wire_numbers = set()

            for base_wire_number, signals in wire_number_groups.items():
                for i, signal_info in enumerate(signals):
                    signal_name = signal_info['signal_name']
                    net_segment_ids = signal_info['net_segment_ids']

                    # Generate unique wire number based on position in sorted group
                    if i == 0:
                        # First signal at this position gets the base wire number
                        unique_wire_number = base_wire_number
                    else:
                        # Subsequent signals get letter suffixes
                        letter = chr(ord('A') + i - 1)
                        unique_wire_number = f"{base_wire_number}.{letter}"

                    used_wire_numbers.add(unique_wire_number)

                    self.logger.info(f"Signal '{signal_name}' assigned unique wire number: {unique_wire_number} (X={signal_info['x_coord']}, Y={signal_info['y_coord']}, {len(net_segment_ids)} net segments)")

                    # Set wire number for all net segments in this signal
                    for net_segment_id in net_segment_ids:
                        try:
                            self.net_segment.SetId(net_segment_id)
                            self.net_segment.SetAttributeValue("Wire number", unique_wire_number)
                            updated_count += 1
                            self.logger.debug(f"Set wire number '{unique_wire_number}' for net segment {net_segment_id}")

                        except Exception as e:
                            self.logger.error(f"Error setting wire number for net segment {net_segment_id}: {e}")

            self.logger.info(f"Successfully updated wire numbers for {updated_count} net segments")
            self.logger.info(f"Total signals processed: {len(signal_data)}")
            self.logger.info(f"Total unique wire numbers assigned: {len(used_wire_numbers)}")

        except Exception as e:
            self.logger.error(f"Error processing connections: {e}")

    def run(self):
        """Main execution method"""
        self.logger.info("Starting wire number assignment process")

        if not self.connect_to_e3():
            self.logger.error("Failed to connect to E3. Make sure E3 is running with a project open.")
            return False

        try:
            self.process_connections()
            self.logger.info("Wire number assignment completed successfully")
            return True

        except Exception as e:
            self.logger.error(f"Error during wire number assignment: {e}")
            return False

        finally:
            # Clean up E3.series objects
            self.app = None
            self.job = None
            self.connection = None
            self.pin = None
            self.sheet = None
            self.signal = None
            self.net = None
            self.net_segment = None
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def run_wire_number_automation(logger=None, e3_pid=None):
    """
    Main function to run wire number automation.

    Args:
        logger: Optional logger instance. If None, creates a default logger.
        e3_pid: Optional E3.series process ID. If None, will prompt user to select.

    Returns:
        bool: True if successful, False otherwise
    """
    if logger is None:
        # Configure default logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('wire_numbering.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        logger = logging.getLogger(__name__)

    assigner = WireNumberAssigner(logger, e3_pid)
    success = assigner.run()

    if success:
        logger.info("Wire number assignment completed successfully!")
        return True
    else:
        logger.error("Wire number assignment failed!")
        return False


def main():
    """Main entry point when run as script"""
    success = run_wire_number_automation()
    return 0 if success else 1


if __name__ == "__main__":
    main()
