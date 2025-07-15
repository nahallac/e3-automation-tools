#!/usr/bin/env python3
"""
E3 Device Designation Automation Library

This module provides functionality to automatically update device designations in E3.series 
projects based on the sheet and grid position of the topmost leftmost pin of the first 
symbol of each device.

The designation format is: {device letter code}{sheet}.{grid}

For devices with conflicting designations, letter suffixes (A, B, C, etc.) are appended
similar to the wire numbering logic.

Special handling for terminal devices:
- Terminal devices (identified using E3 API methods IsTerminal() and IsTerminalBlock()) are completely skipped
- Terminal devices do not get their designations renamed and are ignored by this script
- Fallback to letter code detection (T, TB, X, XT, TERM) if API methods fail

Cable handling:
- Cables are retrieved using e3Job.GetCableIds() and processed as devices
- Cables use the same designation logic as regular devices
- Cable IDs are handled through e3Device objects for designation setting

Author: Jonathan Callahan
Date: 2025-07-07
Updated: 2025-07-08 - Migrated to use e3series PyPI package instead of win32com.client
Updated: 2025-07-11 - Added cable support using e3Job.GetCableIds() and e3Device.IsCable()
Updated: 2025-07-11 - Converted to library module for GUI integration
"""

import logging
import sys
from typing import Dict, List, Tuple, Optional
import e3series
import pythoncom
from .e3_connection_manager import E3ConnectionManager


class DeviceDesignationManager:
    """Manages device and cable designation automation for E3.series projects"""

    def __init__(self, logger=None, e3_app=None):
        self.app = e3_app  # Allow passing existing E3 app instance
        self.job = None
        self.device = None
        self.symbol = None
        self.sheet = None
        self.logger = logger or logging.getLogger(__name__)
        self.owns_connection = e3_app is None  # Track if we created the connection

        
    def connect_to_e3(self):
        """Connect to E3 application with support for multiple instances"""
        try:
            # If we already have an app instance, use it
            if self.app is None:
                # Initialize COM
                pythoncom.CoInitialize()

                # Use the connection manager to handle multiple instances
                connection_manager = E3ConnectionManager(self.logger)
                self.app = connection_manager.connect_to_e3()

                if not self.app:
                    return False
            else:
                self.logger.info("Using provided E3 application instance")

            # Create E3 objects
            self.job = self.app.CreateJobObject()
            self.device = self.job.CreateDeviceObject()
            self.symbol = self.job.CreateSymbolObject()
            self.sheet = self.job.CreateSheetObject()

            self.logger.info("Successfully connected to E3 application using e3series library")
            return True

        except Exception as e:
            self.logger.error(f"Failed to connect to E3 application: {e}")
            return False
    
    def extract_grid_position(self, grid_desc: str) -> str:
        """
        Extract grid position from grid description string.
        
        Args:
            grid_desc: Grid description in format "/sheet.grid" or similar
            
        Returns:
            Grid position string (e.g., "A1", "B2")
        """
        if not grid_desc:
            return ""
            
        # Extract grid part after the last dot
        parts = grid_desc.split('.')
        if len(parts) >= 2:
            return parts[-1]
        
        return grid_desc
    
    def get_device_letter_code(self, device_id: int) -> str:
        """
        Get the device letter code from device attributes or name.
        
        Args:
            device_id: Device ID
            
        Returns:
            Device letter code (e.g., "M", "K", "T")
        """
        try:
            self.device.SetId(device_id)
            
            letercode = self.device.GetComponentAttributeValue("DeviceLetterCode")
            return letercode
            
            
        except Exception as e:
            self.logger.error(f"Error getting device letter code for device {device_id}: {e}")
            return "X"
    

    def generate_device_designation(self, letter_code: str, sheet: str, grid: str) -> str:
        """
        Generate device designation from components.

        Args:
            letter_code: Device letter code (e.g., "M", "K")
            sheet: Sheet page number
            grid: Grid position (e.g., "A1", "B2")

        Returns:
            Device designation string
        """
        return f"{letter_code}{sheet}{grid}"

    def get_first_symbol_info(self, device_id: int) -> Tuple[Optional[str], Optional[str], Optional[int]]:
        """
        Get the sheet, grid position, and first symbol ID using the first symbol returned by E3.

        Args:
            device_id: Device ID

        Returns:
            Tuple of (sheet_assignment, grid_position, first_symbol_id) or (None, None, None) if not found
        """
        try:
            self.device.SetId(device_id)

            # Get all symbols for this device
            symbol_ids_result = self.device.GetSymbolIds()

            if not symbol_ids_result or symbol_ids_result == 0:
                symbol_ids_result = self.device.GetSymbolIds(1)  # 1 = placed symbols only

            if not symbol_ids_result or symbol_ids_result == 0:
                return None, None, None

            # Handle the result format - E3 returns (count, (symbol_id1, symbol_id2, ...))
            symbol_ids = []
            if isinstance(symbol_ids_result, tuple) and len(symbol_ids_result) >= 2:
                symbol_tuple = symbol_ids_result[1]

                if isinstance(symbol_tuple, tuple) and len(symbol_tuple) >= 1:
                    # Process all symbol IDs in the tuple
                    for sid in symbol_tuple:
                        if sid is not None and isinstance(sid, int) and sid > 0:
                            symbol_ids.append(sid)

            if not symbol_ids:
                return None, None, None

            self.logger.debug(f"Device {device_id}: Found {len(symbol_ids)} symbols: {symbol_ids}")

            # Use the FIRST symbol returned by E3 (this is the "first symbol" according to E3's internal order)
            first_symbol_id = symbol_ids[0]

            try:
                self.symbol.SetId(first_symbol_id)
                result = self.symbol.GetSchemaLocation()

                if not result or result == 0:
                    return None, None, None

                if isinstance(result, tuple) and len(result) >= 4:
                    sheet_id, x_pos, y_pos, grid_desc = result[:4]

                    if not isinstance(sheet_id, int) or sheet_id <= 0:
                        return None, None, None

                    # Get sheet assignment
                    self.sheet.SetId(sheet_id)
                    sheet_assignment = self.sheet.GetName()

                    grid_position = self.extract_grid_position(grid_desc)

                    self.logger.debug(f"Device {device_id}: Using first symbol {first_symbol_id} at sheet {sheet_assignment}, grid {grid_position}, pos ({x_pos}, {y_pos})")
                    return sheet_assignment, grid_position, first_symbol_id

            except Exception as e:
                self.logger.debug(f"Error processing first symbol {first_symbol_id} for device {device_id}: {e}")
                return None, None, None

            return None, None, None

        except Exception as e:
            self.logger.error(f"Error getting first symbol info for device {device_id}: {e}")
            return None, None, None

    def is_terminal_device(self, device_id: int) -> bool:
        """
        Check if a device is a terminal using the official E3 API methods.

        This method uses e3Device.IsTerminal() and e3Device.IsTerminalBlock()
        which are the official E3 API methods for identifying terminal devices.

        Args:
            device_id: Device ID

        Returns:
            True if device is a terminal or terminal block, False otherwise
        """
        try:
            self.device.SetId(device_id)

            # Use official E3 API methods to check if device is a terminal
            is_terminal = self.device.IsTerminal()
            is_terminal_block = self.device.IsTerminalBlock()

            # Return True if either method indicates this is a terminal device
            result = (is_terminal == 1) or (is_terminal_block == 1)

            if result:
                self.logger.debug(f"Device {device_id} identified as terminal device (IsTerminal={is_terminal}, IsTerminalBlock={is_terminal_block})")

            return result

        except Exception as e:
            self.logger.error(f"Error checking if device {device_id} is terminal using E3 API: {e}")
            # Fallback to letter code method if API calls fail
            try:
                letter_code = self.get_device_letter_code(device_id)
                # Common terminal letter codes as fallback
                terminal_codes = ['T', 'TB', 'X', 'XT', 'TERM']
                fallback_result = letter_code in terminal_codes
                if fallback_result:
                    self.logger.warning(f"Device {device_id} identified as terminal using fallback letter code method")
                return fallback_result
            except Exception as e2:
                self.logger.error(f"Error in fallback terminal detection for device {device_id}: {e2}")
                return False

    def assign_suffix_for_conflicts(self, designations: Dict[str, List[int]], device_data: Dict) -> Dict[int, str]:
        """
        Assign letter suffixes for conflicting device designations.

        Args:
            designations: Dictionary mapping base designation to list of device IDs
            device_data: Dictionary mapping device ID to device information including first_symbol_id

        Returns:
            Dictionary mapping device ID to final designation with suffix
        """
        final_designations = {}

        for base_designation, device_ids in designations.items():
            if len(device_ids) == 1:
                # No conflict, use base designation
                final_designations[device_ids[0]] = base_designation
            else:
                # Multiple devices with same base designation, sort by first symbol order and add suffixes
                self.logger.info(f"Found {len(device_ids)} devices with designation '{base_designation}', adding suffixes")

                # Sort devices by their device ID for consistent ordering
                device_orders = []
                for device_id in device_ids:
                    try:
                        # Use the first symbol ID that was already identified for this device
                        if device_id in device_data and 'first_symbol_id' in device_data[device_id]:
                            first_symbol_id = device_data[device_id]['first_symbol_id']

                            try:
                                self.symbol.SetId(first_symbol_id)
                                location = self.symbol.GetSchemaLocation()
                                if len(location) >= 3:
                                    x_pos = location[1]
                                    self.logger.debug(f"Device {device_id} first symbol {first_symbol_id}: x_pos={x_pos}")
                                    device_orders.append((device_id, x_pos, device_id))
                                else:
                                    self.logger.debug(f"Could not get location for first symbol {first_symbol_id} of device {device_id}")
                                    device_orders.append((device_id, 0.0, device_id))
                            except Exception as e:
                                self.logger.debug(f"Could not get position for first symbol {first_symbol_id} of device {device_id}: {e}")
                                device_orders.append((device_id, 0.0, device_id))
                        else:
                            self.logger.debug(f"No first symbol ID found for device {device_id}")
                            device_orders.append((device_id, 0.0, device_id))
                    except Exception as e:
                        self.logger.debug(f"Could not get position for device {device_id}: {e}")
                        # If we can't get position, add with default values
                        device_orders.append((device_id, 0.0, device_id))

                # Sort by device ID for consistent ordering
                device_orders.sort(key=lambda x: x[0])

                # Assign suffixes - first device keeps original designation, others get suffixes
                for i, (_, _, device_id) in enumerate(device_orders):
                    if i == 0:
                        # First device keeps the original designation (no suffix)
                        final_designations[device_id] = base_designation
                        self.logger.info(f"Device {device_id} assigned designation: {base_designation} (first device, no suffix)")
                    else:
                        # Subsequent devices get suffixes starting with A
                        suffix = chr(ord('A') + i - 1)  # i-1 so second device gets 'A', third gets 'B', etc.
                        final_designation = f"{base_designation}.{suffix}"
                        final_designations[device_id] = final_designation
                        self.logger.info(f"Device {device_id} assigned designation: {final_designation}")

        return final_designations

    def update_device_designation(self, device_id: int, designation: str) -> bool:
        """
        Update the device designation attribute.

        Args:
            device_id: Device ID
            designation: New designation

        Returns:
            True if successful, False otherwise
        """
        try:
            self.device.SetId(device_id)

            # Use SetName() to set the device designation
            result = self.device.SetName(designation)
            if result > 0:  # Success
                self.logger.info(f"Updated device {device_id} designation to '{designation}' using SetName()")
                return True
            else:
                self.logger.warning(f"SetName() failed for device {device_id}, result: {result}")
                return False

        except Exception as e:
            self.logger.error(f"Error updating designation for device {device_id}: {e}")
            return False

    def get_all_device_and_cable_ids(self):
        """
        Get all device IDs and cable IDs from the project.

        Returns:
            List of all device and cable IDs
        """
        all_devices = []

        # Get regular devices
        try:
            device_ids_result = self.job.GetAllDeviceIds()
            if device_ids_result:
                if isinstance(device_ids_result, tuple) and len(device_ids_result) >= 2:
                    count = device_ids_result[0]
                    device_ids = device_ids_result[1]
                    self.logger.info(f"E3 reports {count} regular devices")

                    if isinstance(device_ids, tuple):
                        # Filter out None values
                        regular_devices = [did for did in device_ids if did is not None]
                        all_devices.extend(regular_devices)
                    else:
                        if device_ids is not None:
                            all_devices.append(device_ids)
                else:
                    self.logger.warning(f"Unexpected device IDs format: {type(device_ids_result)}")
        except Exception as e:
            self.logger.error(f"Error getting regular device IDs: {e}")

        # Get cables (but don't add them if they're already in the device list)
        try:
            cable_ids_result = self.job.GetCableIds()
            if cable_ids_result:
                if isinstance(cable_ids_result, tuple) and len(cable_ids_result) >= 2:
                    count = cable_ids_result[0]
                    cable_ids = cable_ids_result[1]
                    self.logger.info(f"E3 reports {count} cables")

                    if isinstance(cable_ids, tuple):
                        # Filter out None values and duplicates
                        cables = [cid for cid in cable_ids if cid is not None and cid not in all_devices]
                        all_devices.extend(cables)
                    else:
                        if cable_ids is not None and cable_ids not in all_devices:
                            all_devices.append(cable_ids)
                else:
                    self.logger.warning(f"Unexpected cable IDs format: {type(cable_ids_result)}")
        except Exception as e:
            self.logger.error(f"Error getting cable IDs: {e}")

        # Remove any remaining duplicates
        all_devices = list(set(all_devices))

        return all_devices

    def is_cable_device(self, device_id: int) -> bool:
        """
        Check if a device is a cable using the E3 API.

        Args:
            device_id: Device ID

        Returns:
            True if device is a cable, False otherwise
        """
        try:
            self.device.SetId(device_id)
            result = self.device.IsCable()
            return result == 1
        except Exception as e:
            self.logger.debug(f"Error checking if device {device_id} is cable: {e}")
            return False

    def get_cable_position_info(self, device_id: int) -> Tuple[Optional[str], Optional[str], Optional[int]]:
        """
        Get position information for a cable using its first core/connection.

        Args:
            device_id: Cable device ID

        Returns:
            Tuple of (sheet_assignment, grid_position, core_id) or (None, None, None) if not found
        """
        try:
            self.device.SetId(device_id)

            # Get cores/conductors from the cable
            core_ids_result = self.device.GetCoreIds()

            if not core_ids_result:
                self.logger.debug(f"No cores found for cable {device_id}")
                return None, None, None

            # Handle the result format - E3 returns (count, (core_id1, core_id2, ...))
            core_ids = []
            if isinstance(core_ids_result, tuple) and len(core_ids_result) >= 2:
                core_tuple = core_ids_result[1]

                if isinstance(core_tuple, tuple) and len(core_tuple) >= 1:
                    # Process all core IDs in the tuple
                    for cid in core_tuple:
                        if cid is not None and isinstance(cid, int) and cid > 0:
                            core_ids.append(cid)

            if not core_ids:
                self.logger.debug(f"No valid cores found for cable {device_id}")
                return None, None, None

            # Try to get position from the first core
            for core_id in core_ids:
                try:
                    # Note: This would need a pin object, but for now we'll return basic info
                    # In a real implementation, you'd need to create a pin object and use it
                    self.logger.debug(f"Cable {device_id} core {core_id} found")
                    return "Unknown", None, core_id

                except Exception as e:
                    self.logger.debug(f"Error getting position for cable {device_id} core {core_id}: {e}")
                    continue

            self.logger.debug(f"Could not determine position for cable {device_id}")
            return None, None, None

        except Exception as e:
            self.logger.error(f"Error getting cable position info for device {device_id}: {e}")
            return None, None, None

    def process_devices(self):
        """Process all devices and cables in the project"""
        try:
            # Get all device IDs (including cables)
            actual_devices = self.get_all_device_and_cable_ids()

            if not actual_devices:
                self.logger.warning("No devices or cables found in project")
                return

            self.logger.info(f"Processing {len(actual_devices)} devices and cables")

            # Collect device data
            device_data = {}
            designations = {}  # base_designation -> [device_ids]
            devices_with_symbols = 0
            devices_without_symbols = 0
            terminal_devices_skipped = 0
            non_terminal_devices = []

            # First pass: separate terminals from other devices and identify cables
            cable_count = 0
            for device_id in actual_devices:
                try:
                    # Check if this is a terminal device - skip it completely
                    if self.is_terminal_device(device_id):
                        terminal_devices_skipped += 1
                        self.logger.info(f"Device {device_id} identified as terminal device - skipping")
                        continue

                    # Check if this is a cable
                    if self.is_cable_device(device_id):
                        cable_count += 1
                        self.logger.info(f"Device {device_id} identified as cable")
                        # Cables are processed as non-terminal devices for designation purposes
                        non_terminal_devices.append(device_id)
                    else:
                        non_terminal_devices.append(device_id)
                except Exception as e:
                    self.logger.error(f"Error checking device type for {device_id}: {e}")
                    non_terminal_devices.append(device_id)  # Default to non-terminal

            self.logger.info(f"Skipped {terminal_devices_skipped} terminal devices, processing {len(non_terminal_devices)} non-terminal devices (including {cable_count} cables)")

            # Process non-terminal devices for designation updates
            cable_counter = 1  # Sequential counter for cables
            for device_id in non_terminal_devices:
                try:
                    # Check if this is a cable - handle differently
                    if self.is_cable_device(device_id):
                        # Get cable's default letter code (instead of hardcoding "W")
                        letter_code = self.get_device_letter_code(device_id)

                        # For cables, use sequential numbering with their default letter code
                        cable_designation = f"{letter_code}{cable_counter:03d}"  # e.g., CBL001, CBL002, etc.
                        cable_counter += 1

                        # Get position info for logging purposes
                        sheet, grid, _ = self.get_cable_position_info(device_id)

                        # Store cable data
                        device_data[device_id] = {
                            'letter_code': letter_code,
                            'sheet': sheet or 'Unknown',
                            'grid': grid or 'Unknown',
                            'base_designation': cable_designation,
                            'first_symbol_id': None  # Cables don't have symbols
                        }

                        # Track designations for conflict resolution (though cables shouldn't conflict)
                        if cable_designation not in designations:
                            designations[cable_designation] = []
                        designations[cable_designation].append(device_id)

                        self.logger.info(f"Cable {device_id}: {letter_code} at sheet {sheet or 'Unknown'}, grid {grid or 'Unknown'} -> {cable_designation}")
                        devices_with_symbols += 1  # Count as processed
                    else:
                        # Regular device processing
                        # Get device letter code
                        letter_code = self.get_device_letter_code(device_id)

                        # Get position and first symbol ID of topmost leftmost symbol
                        sheet, grid, first_symbol_id = self.get_first_symbol_info(device_id)

                        if sheet and grid and first_symbol_id:
                            # Generate base designation
                            base_designation = self.generate_device_designation(letter_code, sheet, grid)

                            # Store device data including the first symbol ID
                            device_data[device_id] = {
                                'letter_code': letter_code,
                                'sheet': sheet,
                                'grid': grid,
                                'base_designation': base_designation,
                                'first_symbol_id': first_symbol_id
                            }

                            # Track designations for conflict resolution
                            if base_designation not in designations:
                                designations[base_designation] = []
                            designations[base_designation].append(device_id)

                            self.logger.info(f"Device {device_id}: {letter_code} at sheet {sheet}, grid {grid} -> {base_designation}")
                            devices_with_symbols += 1
                        else:
                            self.logger.info(f"Could not determine position for device {device_id} (sheet={sheet}, grid={grid}, symbol_id={first_symbol_id}), counting as without symbols")
                            devices_without_symbols += 1

                except Exception as e:
                    self.logger.error(f"Error processing device {device_id}: {e}")
                    devices_without_symbols += 1
                    continue

            self.logger.info(f"Found {devices_with_symbols} non-terminal devices with placed symbols, {devices_without_symbols} devices without placed symbols")

            # Resolve conflicts and assign final designations for non-terminal devices
            final_designations = self.assign_suffix_for_conflicts(designations, device_data)

            # Update device designations for non-terminal devices
            designation_success_count = 0
            for device_id, final_designation in final_designations.items():
                if self.update_device_designation(device_id, final_designation):
                    designation_success_count += 1

            self.logger.info(f"Successfully updated {designation_success_count} out of {len(final_designations)} device designations")
            self.logger.info(f"Skipped {terminal_devices_skipped} terminal devices (designations unchanged)")

        except Exception as e:
            self.logger.error(f"Error in process_devices: {e}")

    def run(self):
        """Main execution method"""
        try:
            self.logger.info("Starting E3 Device Designation Automation")

            if not self.connect_to_e3():
                return False

            self.process_devices()

            self.logger.info("Device designation automation completed")
            return True

        except Exception as e:
            self.logger.error(f"Error in main execution: {e}")
            return False
        finally:
            # Clean up E3.series objects (but don't clean up shared app instance)
            if self.owns_connection:
                self.app = None
            self.job = None
            self.device = None
            self.symbol = None
            self.sheet = None
            # Only uninitialize COM if we created the connection
            if self.owns_connection:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass


def run_device_designation_automation(logger=None, e3_app=None):
    """
    Main function to run device designation automation.

    Args:
        logger: Optional logger instance. If None, creates a default logger.
        e3_app: Optional existing E3 application instance to reuse.

    Returns:
        bool: True if successful, False otherwise
    """
    if logger is None:
        # Configure default logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('device_designation.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        logger = logging.getLogger(__name__)

    try:
        manager = DeviceDesignationManager(logger, e3_app)
        success = manager.run()

        if success:
            logger.info("Device designation automation completed successfully!")
            return True
        else:
            logger.error("Device designation automation failed!")
            return False

    except Exception as e:
        logger.error(f"Unexpected error in device designation automation: {e}")
        return False


def main():
    """Main entry point when run as script"""
    success = run_device_designation_automation()
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
