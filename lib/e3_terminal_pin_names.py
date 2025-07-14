#!/usr/bin/env python3
"""
E3 Terminal Pin Name Automation Library

This module provides functionality to automatically set terminal pin names to match 
the wire numbers of their connected net segments. It processes all terminal devices 
in the E3.series project and updates each pin's name to match the "Wire number" 
attribute of the connected net segment.

Key Features:
- Identifies terminal devices using E3 API methods IsTerminal() and IsTerminalBlock()
- Gets all pins from terminal devices
- Retrieves wire numbers from connected net segments
- Sets pin names to match wire numbers
- Comprehensive logging and error handling
- Uses the modern e3series PyPI library

Author: E3 Automation Tools
Date: January 2025
Updated: 2025-07-11 - Converted to library module for GUI integration
"""

import logging
import sys
from collections import defaultdict
import e3series
import pythoncom

class TerminalPinNameSetter:
    def __init__(self, logger=None):
        self.app = None
        self.job = None
        self.device = None
        self.pin = None
        self.net_segment = None
        self.connection = None
        self.logger = logger or logging.getLogger(__name__)
        
    def connect_to_e3(self):
        """Connect to the open E3 application"""
        try:
            # Initialize COM
            pythoncom.CoInitialize()

            # Connect to the active E3.series application
            self.app = e3series.Application()
            self.job = self.app.CreateJobObject()
            self.device = self.job.CreateDeviceObject()
            self.pin = self.job.CreatePinObject()
            self.net_segment = self.job.CreateNetSegmentObject()
            self.connection = self.job.CreateConnectionObject()
            self.logger.info("Successfully connected to E3 application using e3series library")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect to E3: {e}")
            return False
    
    def is_terminal_device(self, device_id: int) -> bool:
        """
        Check if a device is a terminal device using E3 API methods.
        
        Args:
            device_id: Device ID to check
            
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
                device_name = self.device.GetName()
                self.logger.debug(f"Device {device_name} ({device_id}) identified as terminal device (IsTerminal={is_terminal}, IsTerminalBlock={is_terminal_block})")
            
            return result
            
        except Exception as e:
            self.logger.error(f"Error checking if device {device_id} is terminal: {e}")
            return False
    
    def get_all_terminal_devices(self):
        """Get all terminal device IDs in the project"""
        try:
            # Get all device IDs
            device_ids_result = self.job.GetAllDeviceIds()
            
            if not device_ids_result or len(device_ids_result) < 2:
                self.logger.warning("No devices found in project")
                return []
            
            device_count = device_ids_result[0]
            device_ids = device_ids_result[1]
            
            if device_count == 0:
                self.logger.warning("No devices found in project")
                return []
            
            # Convert to list if single device
            if not isinstance(device_ids, tuple):
                device_ids = [device_ids] if device_ids is not None else []
            else:
                device_ids = [did for did in device_ids if did is not None]
            
            # Filter for terminal devices only
            terminal_devices = []
            for device_id in device_ids:
                if self.is_terminal_device(device_id):
                    terminal_devices.append(device_id)
            
            self.logger.info(f"Found {len(terminal_devices)} terminal devices out of {device_count} total devices")
            return terminal_devices
            
        except Exception as e:
            self.logger.error(f"Error getting terminal devices: {e}")
            return []
    
    def get_device_pins(self, device_id: int):
        """Get all pin IDs for a device"""
        try:
            self.device.SetId(device_id)
            pin_ids_result = self.device.GetPinIds()
            
            if not pin_ids_result or len(pin_ids_result) < 2:
                return []
            
            pin_count = pin_ids_result[0]
            pin_ids = pin_ids_result[1]
            
            if pin_count == 0:
                return []
            
            # Convert to list if single pin
            if not isinstance(pin_ids, tuple):
                pin_ids = [pin_ids] if pin_ids is not None else []
            else:
                pin_ids = [pid for pid in pin_ids if pid is not None]
            
            return pin_ids
            
        except Exception as e:
            self.logger.error(f"Error getting pins for device {device_id}: {e}")
            return []
    
    def get_pin_net_segments(self, pin_id: int):
        """Get net segment IDs connected to a pin"""
        try:
            self.pin.SetId(pin_id)
            net_segment_ids_result = self.pin.GetNetSegmentIds()
            
            if not net_segment_ids_result or len(net_segment_ids_result) < 2:
                return []
            
            net_segment_count = net_segment_ids_result[0]
            net_segment_ids = net_segment_ids_result[1]
            
            if net_segment_count == 0:
                return []
            
            # Convert to list if single net segment
            if not isinstance(net_segment_ids, tuple):
                net_segment_ids = [net_segment_ids] if net_segment_ids is not None else []
            else:
                net_segment_ids = [nsid for nsid in net_segment_ids if nsid is not None]
            
            return net_segment_ids
            
        except Exception as e:
            self.logger.error(f"Error getting net segments for pin {pin_id}: {e}")
            return []
    
    def get_wire_number_from_net_segment(self, net_segment_id: int):
        """Get wire number attribute from a net segment"""
        try:
            self.net_segment.SetId(net_segment_id)
            wire_number = self.net_segment.GetAttributeValue("Wire number")
            
            # Check if wire number is empty or None
            if not wire_number or wire_number.strip() == "":
                return None
            
            return wire_number.strip()
            
        except Exception as e:
            self.logger.error(f"Error getting wire number from net segment {net_segment_id}: {e}")
            return None
    
    def set_pin_name(self, pin_id: int, new_name: str):
        """Set the name of a pin"""
        try:
            self.pin.SetId(pin_id)
            old_name = self.pin.GetName()
            
            # Only update if the name is different
            if old_name != new_name:
                result = self.pin.SetName(new_name)
                if result == 1:  # Success
                    self.logger.info(f"Updated pin {pin_id}: '{old_name}' -> '{new_name}'")
                    return True
                else:
                    self.logger.warning(f"Failed to set name for pin {pin_id}: SetName() returned {result}")
                    return False
            else:
                self.logger.debug(f"Pin {pin_id} already has correct name: '{new_name}'")
                return True
                
        except Exception as e:
            self.logger.error(f"Error setting name for pin {pin_id}: {e}")
            return False
    
    def process_terminal_pin(self, device_id: int, pin_id: int):
        """Process a single terminal pin to set its name to the wire number"""
        try:
            # Get net segments connected to this pin
            net_segment_ids = self.get_pin_net_segments(pin_id)
            
            if not net_segment_ids:
                self.logger.debug(f"Pin {pin_id} has no connected net segments")
                return False
            
            # Try to get wire number from the first net segment
            # In most cases, terminal pins should only have one net segment
            wire_number = None
            for net_segment_id in net_segment_ids:
                wire_number = self.get_wire_number_from_net_segment(net_segment_id)
                if wire_number:
                    break
            
            if not wire_number:
                self.logger.debug(f"Pin {pin_id} has no wire number in connected net segments")
                return False
            
            # Set the pin name to the wire number
            success = self.set_pin_name(pin_id, wire_number)
            return success
            
        except Exception as e:
            self.logger.error(f"Error processing terminal pin {pin_id}: {e}")
            return False

    def process_all_terminal_pins(self):
        """Process all terminal pins in the project"""
        try:
            terminal_devices = self.get_all_terminal_devices()

            if not terminal_devices:
                self.logger.warning("No terminal devices found in project")
                return

            total_pins_processed = 0
            total_pins_updated = 0

            for device_id in terminal_devices:
                try:
                    self.device.SetId(device_id)
                    device_name = self.device.GetName()

                    pin_ids = self.get_device_pins(device_id)

                    if not pin_ids:
                        self.logger.debug(f"Terminal device {device_name} ({device_id}) has no pins")
                        continue

                    self.logger.info(f"Processing terminal device {device_name} ({device_id}) with {len(pin_ids)} pins")

                    device_pins_updated = 0
                    for pin_id in pin_ids:
                        total_pins_processed += 1
                        if self.process_terminal_pin(device_id, pin_id):
                            total_pins_updated += 1
                            device_pins_updated += 1

                    self.logger.info(f"Updated {device_pins_updated}/{len(pin_ids)} pins for device {device_name}")

                except Exception as e:
                    self.logger.error(f"Error processing terminal device {device_id}: {e}")
                    continue

            self.logger.info(f"Terminal pin processing complete: {total_pins_updated}/{total_pins_processed} pins updated")

        except Exception as e:
            self.logger.error(f"Error in process_all_terminal_pins: {e}")

    def run(self):
        """Main execution method"""
        try:
            self.logger.info("Starting E3 Terminal Pin Name Automation")

            if not self.connect_to_e3():
                self.logger.error("Failed to connect to E3 application")
                return False

            self.process_all_terminal_pins()

            self.logger.info("E3 Terminal Pin Name Automation completed successfully")
            return True

        except Exception as e:
            self.logger.error(f"Error in main execution: {e}")
            return False
        finally:
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def run_terminal_pin_name_automation(logger=None):
    """
    Main function to run terminal pin name automation.

    Args:
        logger: Optional logger instance. If None, creates a default logger.

    Returns:
        bool: True if successful, False otherwise
    """
    if logger is None:
        # Configure default logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('terminal_pin_names.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        logger = logging.getLogger(__name__)

    setter = TerminalPinNameSetter(logger)
    success = setter.run()

    if success:
        logger.info("Terminal pin name automation completed successfully!")
        return True
    else:
        logger.error("Terminal pin name automation failed!")
        return False


def main():
    """Main entry point when run as script"""
    success = run_terminal_pin_name_automation()
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
