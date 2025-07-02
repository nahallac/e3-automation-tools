#!/usr/bin/env python3
"""
E3 Wire Number Assignment Script

This script reads through the current open E3 project and for every connection,
calculates wire numbers based on page number and grid/ladder position, then
sets the "Wire number" attribute on the corresponding net segments.
The wire number format is: {page_number}{grid_position}
The end that results in the lowest wire number is used.
Each signal gets a unique wire number, and ALL segments and connections
belonging to the same signal receive the same wire number. When multiple
signals share the same base wire number (same grid position), letter
suffixes (.A, .B, .C, etc.) are assigned based on left-to-right position.

Author: Jonathan Callahan
Date: 2025-01-01
"""

import win32com.client
import logging
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('wire_numbering.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

class WireNumberAssigner:
    def __init__(self):
        self.app = None
        self.job = None
        self.connection = None
        self.pin = None
        self.sheet = None
        self.signal = None
        self.net = None
        self.net_segment = None

    def connect_to_e3(self):
        """Connect to the open E3 application"""
        try:
            self.app = win32com.client.GetActiveObject("CT.Application")
            self.job = self.app.CreateJobObject()
            self.connection = self.job.CreateConnectionObject()
            self.pin = self.job.CreatePinObject()
            self.sheet = self.job.CreateSheetObject()
            self.signal = self.job.CreateSignalObject()
            self.net = self.job.CreateNetObject()
            self.net_segment = self.job.CreateNetSegmentObject()
            logging.info("Successfully connected to E3 application")
            return True
        except Exception as e:
            logging.error(f"Failed to connect to E3: {e}")
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
                    logging.warning(f"Unexpected GetSchemaLocation result format for pin {pin_id}: {result}")
                    return None, None, None, None, None

            except Exception as e:
                logging.error(f"Error calling GetSchemaLocation for pin {pin_id}: {e}")
                return None, None, None, None, None

            if not sheet_id or sheet_id <= 0:
                logging.debug(f"Pin {pin_id} has no valid schema location (sheet_id: {sheet_id})")
                return None, None, None, None, None


            try:
                self.sheet.SetId(sheet_id)
                page_number = self.sheet.GetName()
            except Exception as e:
                logging.error(f"Error getting sheet assignment for sheet {sheet_id}: {e}")
                page_number = "UNKNOWN"

            # Extract grid position from grid_desc or use column/row
            grid_position = self.extract_grid_position(grid_desc, column, row)

            logging.debug(f"Pin {pin_id}: Sheet {sheet_id}, Page {page_number}, Grid {grid_position}, X={x_coord}, Y={y_coord}")

            return page_number, grid_position, sheet_id, x_coord, y_coord

        except Exception as e:
            logging.error(f"Error getting pin location for pin {pin_id}: {e}")
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
            logging.error(f"Error extracting grid position: {e}")
            return "UNKNOWN"

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
            logging.error(f"Error calculating wire number: {e}")
            return "ERROR"

    def get_connection_wire_numbers_and_positions(self, connection_id):
        """Get potential wire numbers and positions for both ends of a connection"""
        try:
            self.connection.SetId(connection_id)

            # Get pin IDs for this connection
            pin_ids_result = self.connection.GetPinIds()
            if not pin_ids_result:
                logging.warning(f"Connection {connection_id} has no pins")
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
                logging.warning(f"Unexpected pin IDs format for connection {connection_id}: {type(pin_ids_result)}")
                return []

            logging.debug(f"Connection {connection_id} has {len(actual_pin_ids)} valid pins")

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
                    logging.debug(f"Pin {pin_id}: Page {page_number}, Grid {grid_position} -> Wire {wire_number}, X={x_coord}, Y={y_coord}")

            return wire_data

        except Exception as e:
            logging.error(f"Error getting wire numbers for connection {connection_id}: {e}")
            return []

    def get_connection_wire_numbers(self, connection_id):
        """Get potential wire numbers for both ends of a connection (backward compatibility)"""
        wire_data = self.get_connection_wire_numbers_and_positions(connection_id)
        return [data['wire_number'] for data in wire_data]

    def get_lowest_wire_number(self, wire_numbers):
        """Get the lowest wire number from a list"""
        if not wire_numbers:
            return None

        # Sort wire numbers to get the lowest one
        # This will sort lexicographically, which should work for most cases
        sorted_numbers = sorted(wire_numbers)
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
                logging.warning(f"Unexpected net segment IDs format for connection {connection_id}: {type(net_segment_ids_result)}")

            return actual_net_segments

        except Exception as e:
            logging.error(f"Error getting net segments for connection {connection_id}: {e}")
            return []

    def process_connections(self):
        """Process all connections in the project"""
        try:
            # Get all connection IDs
            connection_ids_result = self.job.GetAllConnectionIds()
            if not connection_ids_result:
                logging.warning("No connections found in project")
                return

            # E3 API returns (count, tuple_of_ids)
            actual_connections = []
            if isinstance(connection_ids_result, tuple) and len(connection_ids_result) >= 2:
                count = connection_ids_result[0]
                connection_ids = connection_ids_result[1]

                logging.info(f"E3 reports {count} connections")

                if isinstance(connection_ids, tuple):
                    # Filter out None values
                    actual_connections = [cid for cid in connection_ids if cid is not None]
                else:
                    actual_connections = [connection_ids] if connection_ids is not None else []
            else:
                logging.warning(f"Unexpected connection IDs format: {type(connection_ids_result)}")
                return

            logging.info(f"Found {len(actual_connections)} valid connections to process")

            # Group connections by signal name first
            from collections import defaultdict
            signal_connections = defaultdict(list)

            # First pass: group all connections by signal name
            for conn_id in actual_connections:
                if conn_id is None:
                    continue

                try:
                    self.connection.SetId(conn_id)
                    signal_name = self.connection.GetSignalName()

                    if signal_name:  # Only process connections with valid signal names
                        signal_connections[signal_name].append(conn_id)
                        logging.debug(f"Connection {conn_id} belongs to signal '{signal_name}'")

                except Exception as e:
                    logging.error(f"Error getting signal name for connection {conn_id}: {e}")

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
                        all_wire_data.sort(key=lambda x: x['wire_number'])
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

                        logging.debug(f"Signal '{signal_name}' -> Base Wire: {lowest_wire_data['wire_number']}, X={lowest_wire_data['x_coord']}, Connections: {len(connection_ids)}, Net Segments: {len(set(all_net_segments))}")
                    else:
                        logging.warning(f"Could not calculate wire number for signal '{signal_name}'")

                except Exception as e:
                    logging.error(f"Error processing signal '{signal_name}': {e}")

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

                    logging.info(f"Signal '{signal_name}' assigned unique wire number: {unique_wire_number} (X={signal_info['x_coord']}, Y={signal_info['y_coord']}, {len(net_segment_ids)} net segments)")

                    # Set wire number for all net segments in this signal
                    for net_segment_id in net_segment_ids:
                        try:
                            self.net_segment.SetId(net_segment_id)
                            self.net_segment.SetAttributeValue("Wire number", unique_wire_number)
                            updated_count += 1
                            logging.debug(f"Set wire number '{unique_wire_number}' for net segment {net_segment_id}")

                        except Exception as e:
                            logging.error(f"Error setting wire number for net segment {net_segment_id}: {e}")

            logging.info(f"Successfully updated wire numbers for {updated_count} net segments")
            logging.info(f"Total signals processed: {len(signal_data)}")
            logging.info(f"Total unique wire numbers assigned: {len(used_wire_numbers)}")

        except Exception as e:
            logging.error(f"Error processing connections: {e}")

    def run(self):
        """Main execution method"""
        logging.info("Starting wire number assignment process")

        if not self.connect_to_e3():
            logging.error("Failed to connect to E3. Make sure E3 is running with a project open.")
            return False

        try:
            self.process_connections()
            logging.info("Wire number assignment completed successfully")
            return True

        except Exception as e:
            logging.error(f"Error during wire number assignment: {e}")
            return False

        finally:
            # Clean up COM objects
            self.app = None
            self.job = None
            self.connection = None
            self.pin = None
            self.sheet = None
            self.signal = None
            self.net = None
            self.net_segment = None

def main():
    """Main function"""
    assigner = WireNumberAssigner()
    success = assigner.run()

    if success:
        print("Wire number assignment completed successfully!")
        print("Check the log file 'wire_numbering.log' for details.")
    else:
        print("Wire number assignment failed. Check the log file for errors.")
        sys.exit(1)

if __name__ == "__main__":
    main()
