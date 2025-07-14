#!/usr/bin/env python3
"""
E3 NA Standards GUI

This application provides a graphical user interface for running E3.series automation scripts.
It includes buttons to run:
1. Device Designation Automation
2. Terminal Pin Name Automation  
3. Wire Number Automation

The GUI provides real-time feedback and logging for each operation.

Author: Jonathan Callahan
Date: 2025-07-11
"""

import os
import sys
import logging
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
import customtkinter as ctk
from datetime import datetime

# Add parent directory to path to allow importing from lib
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from lib.e3_device_designation import run_device_designation_automation
    from lib.e3_terminal_pin_names import run_terminal_pin_name_automation
    from lib.e3_wire_numbering import run_wire_number_automation
    from lib.theme_utils import apply_theme
except ImportError as e:
    print(f"Error importing required modules: {e}")
    print("Make sure all required modules are in the lib folder.")
    sys.exit(1)


class LogHandler(logging.Handler):
    """Custom logging handler to display logs in the GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        """Emit a log record to the text widget"""
        try:
            msg = self.format(record)
            # Schedule the GUI update in the main thread
            self.text_widget.after(0, lambda: self._append_log(msg))
        except Exception:
            self.handleError(record)
            
    def _append_log(self, msg):
        """Append log message to text widget (must be called from main thread)"""
        try:
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        except Exception:
            pass


class E3AutomationGUI(ctk.CTk):
    """Main GUI application for E3 automation tools"""
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("E3 NA Standards Automation")
        self.geometry("900x600")
        self.minsize(800, 500)
        
        # Apply theme
        try:
            apply_theme("red", "dark")
        except:
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")
        
        # Initialize variables
        self.running_operation = False
        
        # Create widgets
        self.create_widgets()
        
        # Setup logging
        self.setup_logging()
        
    def create_widgets(self):
        """Create the GUI widgets"""
        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Header
        self.grid_rowconfigure(1, weight=0)  # Buttons
        self.grid_rowconfigure(2, weight=1)  # Log area
        self.grid_rowconfigure(3, weight=0)  # Status
        
        # Header
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="E3 NA Standards Automation",
            font=("Arial", 24, "bold")
        )
        title_label.pack(pady=20)
        
        description_label = ctk.CTkLabel(
            header_frame,
            text="Run E3.series automation scripts with real-time feedback",
            font=("Arial", 14)
        )
        description_label.pack(pady=(0, 20))
        
        # Buttons frame
        buttons_frame = ctk.CTkFrame(self)
        buttons_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        buttons_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        # Device Designation button
        self.device_designation_btn = ctk.CTkButton(
            buttons_frame,
            text="Set Device\nDesignations",
            command=self.run_device_designation,
            width=200,
            height=60,
            font=("Arial", 14, "bold"),
            fg_color="#C53F3F",
            hover_color="#A02222"
        )
        self.device_designation_btn.grid(row=0, column=0, padx=10, pady=20)
        
        # Terminal Pin Names button
        self.terminal_pin_btn = ctk.CTkButton(
            buttons_frame,
            text="Set Terminal\nPin Names",
            command=self.run_terminal_pin_names,
            width=200,
            height=60,
            font=("Arial", 14, "bold"),
            fg_color="#C53F3F",
            hover_color="#A02222"
        )
        self.terminal_pin_btn.grid(row=0, column=1, padx=10, pady=20)
        
        # Wire Numbers button
        self.wire_numbers_btn = ctk.CTkButton(
            buttons_frame,
            text="Set Wire\nNumbers",
            command=self.run_wire_numbers,
            width=200,
            height=60,
            font=("Arial", 14, "bold"),
            fg_color="#C53F3F",
            hover_color="#A02222"
        )
        self.wire_numbers_btn.grid(row=0, column=2, padx=10, pady=20)

        # Run All button
        self.run_all_btn = ctk.CTkButton(
            buttons_frame,
            text="Run All\n(Wire → Device → Terminal)",
            command=self.run_all_automation,
            width=220,
            height=60,
            font=("Arial", 14, "bold"),
            fg_color="#2AA876",
            hover_color="#1F8A5F"
        )
        self.run_all_btn.grid(row=0, column=3, padx=10, pady=20)
        
        # Log area
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)
        
        log_label = ctk.CTkLabel(
            log_frame,
            text="Operation Log:",
            font=("Arial", 16, "bold"),
            anchor="w"
        )
        log_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        # Create log text widget using tkinter (customtkinter doesn't have a good text widget)
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            state='disabled',
            bg="#2b2b2b",
            fg="#ffffff",
            font=("Consolas", 10),
            insertbackground="#ffffff"
        )
        self.log_text.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        # Status bar
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.status_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Ready - Select an operation to begin",
            anchor="w",
            font=("Arial", 12)
        )
        self.status_label.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        # Clear log button
        self.clear_log_btn = ctk.CTkButton(
            self.status_frame,
            text="Clear Log",
            command=self.clear_log,
            width=100,
            height=30,
            fg_color="#666666",
            hover_color="#555555"
        )
        self.clear_log_btn.grid(row=0, column=1, padx=20, pady=10)
        
    def setup_logging(self):
        """Setup logging to display in the GUI"""
        # Create a logger for the GUI
        self.logger = logging.getLogger("E3AutomationGUI")
        self.logger.setLevel(logging.INFO)
        
        # Create custom handler for the text widget
        self.log_handler = LogHandler(self.log_text)
        self.log_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.log_handler.setFormatter(formatter)
        
        # Add handler to logger
        self.logger.addHandler(self.log_handler)
        
        # Initial log message
        self.logger.info("E3 Automation GUI started")
        
    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.logger.info("Log cleared")
        
    def set_buttons_enabled(self, enabled):
        """Enable or disable all operation buttons"""
        state = "normal" if enabled else "disabled"
        self.device_designation_btn.configure(state=state)
        self.terminal_pin_btn.configure(state=state)
        self.wire_numbers_btn.configure(state=state)
        self.run_all_btn.configure(state=state)
        
    def update_status(self, message, color=None):
        """Update the status label"""
        self.status_label.configure(text=message)
        if color:
            self.status_label.configure(text_color=color)
            
    def run_operation_thread(self, operation_func, operation_name):
        """Run an operation in a separate thread"""
        try:
            self.running_operation = True
            self.set_buttons_enabled(False)
            self.update_status(f"Running {operation_name}...", "#FFA500")
            
            self.logger.info(f"Starting {operation_name}")
            
            # Run the operation with the GUI logger
            success = operation_func(self.logger)
            
            if success:
                self.logger.info(f"{operation_name} completed successfully!")
                self.update_status(f"{operation_name} completed successfully", "#00AA00")
            else:
                self.logger.error(f"{operation_name} failed!")
                self.update_status(f"{operation_name} failed", "#FF0000")
                messagebox.showerror("Error", f"{operation_name} failed! Check the log for details.")
                
        except Exception as e:
            self.logger.error(f"Error during {operation_name}: {e}")
            self.update_status(f"Error during {operation_name}", "#FF0000")
            messagebox.showerror("Error", f"Error during {operation_name}: {e}")
            
        finally:
            self.running_operation = False
            self.set_buttons_enabled(True)
            
    def run_device_designation(self):
        """Run device designation automation"""
        if self.running_operation:
            return
            
        thread = threading.Thread(
            target=self.run_operation_thread,
            args=(run_device_designation_automation, "Device Designation Automation"),
            daemon=True
        )
        thread.start()
        
    def run_terminal_pin_names(self):
        """Run terminal pin name automation"""
        if self.running_operation:
            return
            
        thread = threading.Thread(
            target=self.run_operation_thread,
            args=(run_terminal_pin_name_automation, "Terminal Pin Name Automation"),
            daemon=True
        )
        thread.start()
        
    def run_wire_numbers(self):
        """Run wire number automation"""
        if self.running_operation:
            return
            
        thread = threading.Thread(
            target=self.run_operation_thread,
            args=(run_wire_number_automation, "Wire Number Automation"),
            daemon=True
        )
        thread.start()

    def run_all_operations_thread(self):
        """Run all three operations in sequence: Wire Numbers → Device Designation → Terminal Pins"""
        try:
            self.running_operation = True
            self.set_buttons_enabled(False)
            self.update_status("Running all automation scripts...", "#FFA500")

            operations = [
                (run_wire_number_automation, "Wire Number Automation"),
                (run_device_designation_automation, "Device Designation Automation"),
                (run_terminal_pin_name_automation, "Terminal Pin Name Automation")
            ]

            all_successful = True

            for i, (operation_func, operation_name) in enumerate(operations, 1):
                self.logger.info(f"Starting {operation_name} ({i}/3)")
                self.update_status(f"Running {operation_name} ({i}/3)...", "#FFA500")

                # Run the operation with the GUI logger
                success = operation_func(self.logger)

                if success:
                    self.logger.info(f"{operation_name} completed successfully!")
                else:
                    self.logger.error(f"{operation_name} failed!")
                    all_successful = False
                    break  # Stop on first failure

            if all_successful:
                self.logger.info("All automation scripts completed successfully!")
                self.update_status("All automation scripts completed successfully", "#00AA00")
            else:
                self.logger.error("Automation sequence failed!")
                self.update_status("Automation sequence failed", "#FF0000")
                messagebox.showerror("Error", "Automation sequence failed! Check the log for details.")

        except Exception as e:
            self.logger.error(f"Error during automation sequence: {e}")
            self.update_status("Error during automation sequence", "#FF0000")
            messagebox.showerror("Error", f"Error during automation sequence: {e}")

        finally:
            self.running_operation = False
            self.set_buttons_enabled(True)

    def run_all_automation(self):
        """Run all automation scripts in sequence"""
        if self.running_operation:
            return

        thread = threading.Thread(
            target=self.run_all_operations_thread,
            daemon=True
        )
        thread.start()


def main():
    """Main entry point"""
    try:
        app = E3AutomationGUI()
        app.mainloop()
    except Exception as e:
        print(f"Error starting E3 Automation GUI: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
