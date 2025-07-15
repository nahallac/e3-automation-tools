#!/usr/bin/env python3
"""
E3 Connection Manager

This module provides functionality to connect to E3.series applications with support
for selecting from multiple running instances, similar to the VBA ConnectToE3() function.

Based on the VBA implementation that uses CT.Dispatcher and CT.DispatcherViewer to
handle multiple E3 instances.

Author: Jonathan Callahan
Date: 2025-07-15
"""

import logging
import sys
import pythoncom
import win32com.client
import e3series
from typing import Optional, List, Tuple
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk


class E3InstanceInfo:
    """Information about an E3.series instance"""
    
    def __init__(self, process_id: int, window_title: str, project_name: str = ""):
        self.process_id = process_id
        self.window_title = window_title
        self.project_name = project_name
        
    def __str__(self):
        if self.project_name and self.project_name not in ["", "Unknown Project", "No Project"]:
            return f"PID {self.process_id}: {self.project_name}"
        elif self.project_name == "No Project":
            return f"PID {self.process_id}: (No Project Open)"
        elif self.project_name == "Project Open":
            return f"PID {self.process_id}: (Project Open - Name Unknown)"
        else:
            return f"PID {self.process_id}: E3.series Instance"


class E3InstanceSelector:
    """GUI for selecting an E3.series instance when multiple are running"""

    def __init__(self, instances: List[E3InstanceInfo]):
        self.instances = instances
        self.selected_instance = None
        self.root = None

    def show_selector(self) -> Optional[E3InstanceInfo]:
        """Show the instance selector dialog and return the selected instance"""
        # Set up CustomTkinter theme to match project style
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTkToplevel()
        self.root.title("E3.series Connection Manager")
        self.root.geometry("700x600")
        self.root.resizable(False, False)

        # Set window icon if available
        try:
            # Try to set a window icon (optional)
            pass
        except:
            pass

        # Make window modal and center it
        self.root.transient()
        self.root.grab_set()
        self.root.focus_set()
        self._center_window()

        # Create widgets
        self._create_widgets()

        # Run the dialog
        self.root.mainloop()

        return self.selected_instance
    
    def _center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """Create the dialog widgets using CustomTkinter"""
        # Configure grid layout
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=0)  # Header
        self.root.grid_rowconfigure(1, weight=1)  # List
        self.root.grid_rowconfigure(2, weight=0)  # Buttons

        # Header frame
        header_frame = ctk.CTkFrame(self.root)
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # Title
        title_label = ctk.CTkLabel(
            header_frame,
            text="Multiple E3.series instances detected",
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=(20, 10))

        # Instructions
        instruction_label = ctk.CTkLabel(
            header_frame,
            text="Please select which E3.series instance to connect to:",
            font=("Arial", 12)
        )
        instruction_label.pack(pady=(0, 20))

        # List frame
        list_frame = ctk.CTkFrame(self.root)
        list_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        # List label
        list_label = ctk.CTkLabel(
            list_frame,
            text="Available E3.series Instances:",
            font=("Arial", 14, "bold"),
            anchor="w"
        )
        list_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # Create a custom listbox using CTkScrollableFrame
        self.scrollable_frame = ctk.CTkScrollableFrame(
            list_frame,
            height=350
        )
        self.scrollable_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        # Track selected instance
        self.selected_index = 0 if self.instances else -1
        self.instance_buttons = []
        self.instance_labels = []  # Store label references for color updates

        # Create clickable list items
        for i, instance in enumerate(self.instances):
            # Create button that looks like a list item
            instance_btn = ctk.CTkButton(
                self.scrollable_frame,
                text="",  # We'll set this with custom layout
                command=lambda idx=i: self._select_instance(idx),
                height=80,
                fg_color="transparent",
                hover_color="#404040",
                border_width=1,
                border_color="#404040",
                anchor="w"
            )
            instance_btn.grid(row=i, column=0, padx=5, pady=5, sticky="ew")

            # Create custom content frame inside the button
            content_frame = ctk.CTkFrame(instance_btn, fg_color="transparent")
            content_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            content_frame.grid_columnconfigure(0, weight=1)

            # Process ID and main info
            main_info = f"PID {instance.process_id}"
            main_label = ctk.CTkLabel(
                content_frame,
                text=main_info,
                font=("Arial", 14, "bold"),
                anchor="w"
            )
            main_label.grid(row=0, column=0, padx=20, pady=(12, 4), sticky="ew")

            # Project name or status
            if instance.project_name and instance.project_name not in ["Unknown Project", "No Project", "E3.series Running"]:
                project_text = f"Project: {instance.project_name}"
                text_color = "#FC0303"  # Green for project names
            else:
                status_text = instance.project_name if instance.project_name else "No Project Open"
                project_text = f"Status: {status_text}"
                text_color = "#FFA500"  # Orange for status

            project_label = ctk.CTkLabel(
                content_frame,
                text=project_text,
                font=("Arial", 12),
                anchor="w",
                text_color=text_color
            )
            project_label.grid(row=1, column=0, padx=20, pady=(0, 12), sticky="ew")

            # Store button and label references
            self.instance_buttons.append(instance_btn)
            self.instance_labels.append({
                'main': main_label,
                'project': project_label,
                'project_color': text_color
            })

            # Make labels clickable too
            main_label.bind("<Button-1>", lambda e, idx=i: self._select_instance(idx))
            project_label.bind("<Button-1>", lambda e, idx=i: self._select_instance(idx))

        # Select first item by default
        if self.instances:
            self._select_instance(0)

        # Button frame
        button_frame = ctk.CTkFrame(self.root)
        button_frame.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)

        # Connect button
        connect_button = ctk.CTkButton(
            button_frame,
            text="Connect",
            command=self._on_ok,
            width=140,
            height=40,
            font=("Arial", 14, "bold"),
            fg_color="#C53F3F",
            hover_color="#A02222"
        )
        connect_button.grid(row=0, column=0, padx=10, pady=20)

        # Cancel button
        cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self._on_cancel,
            width=140,
            height=40,
            font=("Arial", 14),
            fg_color="#666666",
            hover_color="#555555"
        )
        cancel_button.grid(row=0, column=1, padx=10, pady=20)

        # Bind Enter key to Connect
        self.root.bind("<Return>", lambda e: self._on_ok())
        self.root.bind("<Escape>", lambda e: self._on_cancel())

    def _select_instance(self, index):
        """Select an instance and update visual feedback"""
        if 0 <= index < len(self.instances):
            # Update selected index
            self.selected_index = index

            # Update button appearances and label colors
            for i, (btn, labels) in enumerate(zip(self.instance_buttons, self.instance_labels)):
                if i == index:
                    # Selected item - red background with white text
                    btn.configure(
                        fg_color="#C53F3F",
                        hover_color="#A02222",
                        border_color="#C53F3F"
                    )
                    labels['main'].configure(text_color="white")
                    labels['project'].configure(text_color="white")
                else:
                    # Unselected items - transparent background with original colors
                    btn.configure(
                        fg_color="transparent",
                        hover_color="#404040",
                        border_color="#404040"
                    )
                    labels['main'].configure(text_color="white")  # Default text color
                    labels['project'].configure(text_color=labels['project_color'])  # Original color

    def _on_ok(self):
        """Handle Connect button click"""
        if 0 <= self.selected_index < len(self.instances):
            self.selected_instance = self.instances[self.selected_index]
        else:
            self.selected_instance = None
        self.root.destroy()

    def _on_cancel(self):
        """Handle Cancel button click"""
        self.selected_instance = None
        self.root.destroy()


class E3ConnectionManager:
    """Manager for connecting to E3.series applications"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or logging.getLogger(__name__)
        self.app = None
        
    def connect_to_e3(self) -> Optional[object]:
        """
        Connect to E3.series application with support for multiple instances.
        
        Returns:
            E3 application object if successful, None otherwise
        """
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Check for E3.series instances using process enumeration
            instances = self._get_e3_instances_basic()

            if len(instances) > 1:
                # Multiple instances - show selector
                self.logger.info(f"Found {len(instances)} E3.series instances")
                selector = E3InstanceSelector(instances)
                selected = selector.show_selector()

                if selected is None:
                    self.logger.info("User cancelled E3 instance selection")
                    return None

                self.logger.info(f"User selected: {selected}")
                # Connect to E3 (the e3series library will connect to the active instance)
                self.app = e3series.Application()

            elif len(instances) == 1:
                # Single instance - connect directly
                self.logger.info("Found single E3.series instance, connecting...")
                self.app = e3series.Application()

            else:
                # No instances found
                self.logger.error("No E3.series instances found")
                messagebox.showerror(
                    "E3.series Not Found",
                    "No running E3.series instances found.\n"
                    "Please start E3.series with a project open and try again."
                )
                return None
            
            # Verify connection by creating a job object
            if self.app:
                job = self.app.CreateJobObject()
                if job:
                    self.logger.info("Successfully connected to E3.series application")
                    return self.app
                else:
                    self.logger.error("Failed to create job object - no project may be open")
                    messagebox.showerror(
                        "E3.series Connection Error",
                        "Connected to E3.series but no project appears to be open.\n"
                        "Please open a project in E3.series and try again."
                    )
                    return None
            
        except Exception as e:
            self.logger.error(f"Failed to connect to E3.series: {e}")
            messagebox.showerror(
                "E3.series Connection Error",
                f"Failed to connect to E3.series:\n{str(e)}\n\n"
                "Please ensure E3.series is running with a project open."
            )
            return None
        
        return None

    def _get_e3_instances_via_dispatcher(self) -> List[E3InstanceInfo]:
        """
        Get E3.series instances using CT.Dispatcher (preferred method).

        Returns:
            List of E3InstanceInfo objects
        """
        instances = []

        try:
            # Try to create CT.Dispatcher object
            disp = win32com.client.Dispatch("CT.Dispatcher")

            # Get E3 applications
            result = disp.GetE3Applications()

            if isinstance(result, tuple) and len(result) >= 2:
                count = result[0]
                app_list = result[1]

                self.logger.info(f"CT.Dispatcher found {count} E3.series applications")

                if count > 0:
                    # If we have multiple applications, we can enumerate them
                    if isinstance(app_list, (list, tuple)):
                        for i, app_info in enumerate(app_list):
                            # Extract information about each instance
                            # Note: The exact structure depends on E3.series version
                            instances.append(E3InstanceInfo(
                                process_id=i + 1,  # Use index as pseudo-PID
                                window_title=f"E3.series Instance {i + 1}",
                                project_name=""  # Would need additional API calls to get project name
                            ))
                    else:
                        # Single instance
                        instances.append(E3InstanceInfo(
                            process_id=1,
                            window_title="E3.series Instance",
                            project_name=""
                        ))

        except Exception as e:
            self.logger.debug(f"CT.Dispatcher not available or failed: {e}")

        return instances

    def _get_e3_instances_via_wmi(self) -> List[E3InstanceInfo]:
        """
        Get E3.series instances using WMI (fallback method).

        Returns:
            List of E3InstanceInfo objects
        """
        instances = []

        try:
            import wmi

            # Connect to WMI
            c = wmi.WMI()

            # Query for E3.series processes
            processes = c.query("SELECT * FROM Win32_Process WHERE Name LIKE '%E3%' OR Caption LIKE '%E3.series%'")

            for process in processes:
                # Filter for actual E3.series processes
                if "E3.series" in str(process.Caption) or "E3" in str(process.Name):
                    instances.append(E3InstanceInfo(
                        process_id=process.ProcessId,
                        window_title=str(process.Caption),
                        project_name=""  # Would need additional methods to get project name
                    ))

        except ImportError:
            self.logger.debug("WMI module not available, using basic process enumeration")
            # Fallback to basic process enumeration
            instances = self._get_e3_instances_basic()
        except Exception as e:
            self.logger.debug(f"WMI query failed: {e}")
            instances = self._get_e3_instances_basic()

        return instances

    def _get_e3_instances_basic(self) -> List[E3InstanceInfo]:
        """
        Basic E3.series instance detection using tasklist.

        Returns:
            List of E3InstanceInfo objects with project information
        """
        instances = []

        try:
            import subprocess

            # Use tasklist to find E3.series processes
            result = subprocess.run(
                ['tasklist', '/FO', 'CSV'],
                capture_output=True,
                text=True,
                shell=True
            )

            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                # Skip header line
                for line in lines[1:]:
                    if line.strip() and 'E3.series.exe' in line:
                        # Parse CSV format: "Image Name","PID","Session Name","Session#","Mem Usage"
                        parts = [part.strip('"') for part in line.split('","')]
                        if len(parts) >= 2:
                            try:
                                pid = int(parts[1])

                                # Try to get project information for this instance
                                project_name = self._get_project_name_for_instance(pid)

                                instances.append(E3InstanceInfo(
                                    process_id=pid,
                                    window_title=f"E3.series (PID: {pid})",
                                    project_name=project_name
                                ))
                            except ValueError:
                                continue

        except Exception as e:
            self.logger.debug(f"Basic process enumeration failed: {e}")

        return instances

    def _get_project_name_for_instance(self, pid: int) -> str:
        """
        Try to get the project name for a specific E3.series instance.

        Args:
            pid: Process ID of the E3.series instance

        Returns:
            Project name if found, otherwise a descriptive string
        """
        try:
            # Try to get window titles to identify projects
            import win32gui
            import win32process

            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if window_pid == pid:
                        window_title = win32gui.GetWindowText(hwnd)
                        if window_title:
                            windows.append(window_title)
                return True

            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)

            # Look for E3.series main window and project-related windows
            main_window = None
            project_windows = []

            for window_title in windows:
                if 'E3.series' in window_title or 'E3.DTM' in window_title:
                    if len(window_title) > 15:  # Main window usually has longer title
                        main_window = window_title
                    project_windows.append(window_title)

            # Try to extract project name from main window title
            if main_window:
                self.logger.debug(f"PID {pid} main window: {main_window}")

                # Common E3.series title formats observed:
                # "ProjectName.e3s - E3.DTM Panel Professional - [Sheet info]"
                # "ProjectName.e3s - E3.DTM Panel Professional"
                # "E3.series - ProjectName"
                # "ProjectName - E3.series"

                # Pattern 1: "ProjectName.e3s - E3.DTM Panel Professional..."
                if ".e3s - E3.DTM" in main_window or ".e3s - E3.series" in main_window:
                    # Extract everything before " - E3."
                    if " - E3." in main_window:
                        project_name = main_window.split(" - E3.")[0].strip()
                        if project_name:
                            return project_name

                # Pattern 2: "ProjectName - E3.series"
                elif main_window.endswith(" - E3.series"):
                    project_name = main_window.replace(" - E3.series", "").strip()
                    if project_name:
                        return project_name

                # Pattern 3: "E3.series - ProjectName"
                elif main_window.startswith("E3.series - "):
                    project_name = main_window.replace("E3.series - ", "").strip()
                    # Remove sheet/page information if present
                    if " - " in project_name:
                        project_name = project_name.split(" - ")[0]
                    if project_name and project_name not in ["Untitled", "New"]:
                        return project_name

                # Pattern 4: "E3.series - [ProjectName]"
                elif " - [" in main_window and "]" in main_window:
                    start = main_window.find(" - [") + 4
                    end = main_window.find("]", start)
                    if start > 3 and end > start:
                        project_name = main_window[start:end].strip()
                        if project_name:
                            return project_name

                # Pattern 5: General parsing - look for meaningful project names
                elif " - " in main_window:
                    parts = main_window.split(" - ")
                    for part in parts:
                        part = part.strip()
                        # Look for parts that look like project names
                        if (part and
                            part not in ["E3.series", "E3.DTM Panel Professional", "Untitled", "New", "Document"] and
                            not part.startswith("Sheet") and
                            not part.startswith("Page") and
                            not part.startswith("[") and
                            len(part) > 2):
                            return part

            # Check if any windows suggest a project is open
            has_project_windows = any(
                window for window in windows
                if any(keyword in window.lower() for keyword in ['project', 'sheet', 'page', 'drawing'])
                and 'E3' in window
            )

            if has_project_windows:
                return "Project Open"
            elif windows:
                return "No Project"
            else:
                return "Instance Not Responding"

        except Exception as e:
            self.logger.debug(f"Failed to get project name for PID {pid}: {e}")
            return "Unknown Project"

    def get_project_info(self) -> Optional[Tuple[str, str]]:
        """
        Get information about the currently open project.

        Returns:
            Tuple of (project_name, project_path) if successful, None otherwise
        """
        if not self.app:
            return None

        try:
            job = self.app.CreateJobObject()
            if job:
                # Try different methods to get project info
                project_name = "Unknown Project"
                project_path = "Unknown Path"

                # Try to get project name using different API methods
                try:
                    if hasattr(job, 'GetProjectName'):
                        project_name = job.GetProjectName()
                    elif hasattr(job, 'GetName'):
                        project_name = job.GetName()
                    elif hasattr(self.app, 'GetProjectName'):
                        project_name = self.app.GetProjectName()
                except:
                    pass

                # Try to get project path
                try:
                    if hasattr(job, 'GetProjectPath'):
                        project_path = job.GetProjectPath()
                    elif hasattr(job, 'GetPath'):
                        project_path = job.GetPath()
                    elif hasattr(self.app, 'GetProjectPath'):
                        project_path = self.app.GetProjectPath()
                except:
                    pass

                return (project_name, project_path)
        except Exception as e:
            self.logger.debug(f"Failed to get project info: {e}")

        return None


def create_e3_connection(logger: Optional[logging.Logger] = None) -> Optional[object]:
    """
    Convenience function to create an E3.series connection.

    Args:
        logger: Optional logger instance

    Returns:
        E3 application object if successful, None otherwise
    """
    manager = E3ConnectionManager(logger)
    return manager.connect_to_e3()


def create_shared_e3_connection(logger: Optional[logging.Logger] = None) -> Tuple[Optional[object], Optional[E3ConnectionManager]]:
    """
    Create an E3.series connection and return both the app and manager for reuse.

    Args:
        logger: Optional logger instance

    Returns:
        Tuple of (E3 application object, connection manager) if successful, (None, None) otherwise
    """
    manager = E3ConnectionManager(logger)
    app = manager.connect_to_e3()
    return app, manager
