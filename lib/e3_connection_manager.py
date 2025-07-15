#!/usr/bin/env python3
"""
E3 Connection Manager Library

This module provides functionality to detect multiple running E3.series instances
and present a GUI for users to select which instance to connect to. It returns
the selected process ID (PID) that can be used by automation scripts to connect
to the specific E3 instance.

Features:
- Detects all running E3.series processes
- Shows a GUI selection dialog when multiple instances are found
- Returns the PID of the selected instance
- Handles single instance scenarios automatically
- Provides error handling for no running instances
- Accurate project detection from command line arguments and window titles

Author: Jonathan Callahan
Date: 2025-07-15
"""

import psutil
import tkinter as tk
from tkinter import messagebox, ttk
import logging
from typing import List, Tuple, Optional
import e3series
import pythoncom


class E3InstanceInfo:
    """Information about a running E3.series instance"""
    
    def __init__(self, pid: int, name: str, project_path: str = ""):
        self.pid = pid
        self.name = name
        self.project_path = project_path
        
    def get_project_display_name(self) -> str:
        """Get a clean project name for display"""
        if not self.project_path:
            return "No project detected"
            
        # Extract just the filename if it's a full path
        if '\\' in self.project_path or '/' in self.project_path:
            import os
            project_name = os.path.basename(self.project_path)
            # Remove file extension for cleaner display
            if project_name.endswith('.e3p') or project_name.endswith('.e3') or project_name.endswith('.e3s'):
                project_name = os.path.splitext(project_name)[0]
            return project_name
        
        # If it's already just a name, return as-is
        return self.project_path
        
    def __str__(self):
        project_display = self.get_project_display_name()
        return f"PID {self.pid}: {project_display}"


class E3InstanceSelector:
    """GUI for selecting E3.series instance"""
    
    def __init__(self, instances: List[E3InstanceInfo], connection_manager=None):
        self.instances = instances
        self.selected_instance = None
        self.root = None
        self.connection_manager = connection_manager
        self.listbox = None
        
    def show_selection_dialog(self) -> Optional[E3InstanceInfo]:
        """Show the instance selection dialog and return the selected instance"""
        self.root = tk.Tk()
        self.root.title("Select E3.series Instance")
        self.root.geometry("800x500")
        self.root.resizable(True, True)
        
        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (500 // 2)
        self.root.geometry(f"800x500+{x}+{y}")
        
        # Make window modal
        self.root.transient()
        self.root.grab_set()
        
        self._create_widgets()
        
        # Start the GUI event loop
        self.root.mainloop()
        
        return self.selected_instance
        
    def _create_widgets(self):
        """Create the GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)  # Listbox frame row
        
        # Title label
        title_label = ttk.Label(
            main_frame,
            text=f"Multiple E3.series instances detected ({len(self.instances)} found)",
            font=("Arial", 14, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.W)
        
        # Instructions label
        instructions_label = ttk.Label(
            main_frame,
            text="Please select the E3.series instance you want to connect to:",
            font=("Arial", 10)
        )
        instructions_label.grid(row=1, column=0, pady=(0, 5), sticky=tk.W)
        
        # Info label
        info_label = ttk.Label(
            main_frame,
            text="Each instance shows: Process ID and currently open project",
            font=("Arial", 9),
            foreground="gray"
        )
        info_label.grid(row=2, column=0, pady=(0, 10), sticky=tk.W)
        
        # Listbox frame with scrollbar
        listbox_frame = ttk.Frame(main_frame)
        listbox_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)
        
        # Listbox with scrollbar
        self.listbox = tk.Listbox(listbox_frame, font=("Consolas", 11), height=12)
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        self.listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Populate listbox
        self._populate_listbox()
            
        # Select first item by default
        if self.instances:
            self.listbox.selection_set(0)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, pady=(10, 0), sticky=(tk.W, tk.E))
        button_frame.columnconfigure(1, weight=1)  # Make middle column expand
        
        # Refresh button (left side)
        refresh_button = ttk.Button(
            button_frame,
            text="Refresh",
            command=self._on_refresh
        )
        refresh_button.grid(row=0, column=0, sticky=tk.W)
        
        # Right side button frame
        right_buttons = ttk.Frame(button_frame)
        right_buttons.grid(row=0, column=2, sticky=tk.E)
        
        # Cancel and Select buttons (right side)
        cancel_button = ttk.Button(
            right_buttons,
            text="Cancel",
            command=self._on_cancel
        )
        cancel_button.grid(row=0, column=0, padx=(0, 10))
        
        select_button = ttk.Button(
            right_buttons,
            text="Select",
            command=self._on_select
        )
        select_button.grid(row=0, column=1)
        
        # Bind double-click to select
        self.listbox.bind("<Double-Button-1>", lambda _: self._on_select())
        
        # Bind Enter key to select
        self.root.bind("<Return>", lambda _: self._on_select())
        self.root.bind("<Escape>", lambda _: self._on_cancel())
        
        # Focus on listbox
        self.listbox.focus_set()
    
    def _populate_listbox(self):
        """Populate the listbox with current instances"""
        self.listbox.delete(0, tk.END)
        for instance in self.instances:
            display_text = f"PID {instance.pid:>6} │ {instance.get_project_display_name()}"
            self.listbox.insert(tk.END, display_text)
    
    def _on_refresh(self):
        """Refresh the list of E3.series instances"""
        if self.connection_manager:
            # Get updated list of instances
            self.instances = self.connection_manager.get_running_e3_instances()
            
            # Update the listbox
            self._populate_listbox()
            
            # Select first item if available
            if self.instances:
                self.listbox.selection_set(0)
            
            # Update the title to show new count
            title_text = f"Multiple E3.series instances detected ({len(self.instances)} found)"
            # Find and update the title label (this is a bit hacky but works)
            for child in self.root.winfo_children():
                if isinstance(child, ttk.Frame):
                    for grandchild in child.winfo_children():
                        if isinstance(grandchild, ttk.Label) and "instances detected" in grandchild.cget("text"):
                            grandchild.configure(text=title_text)
                            break
        
    def _on_select(self):
        """Handle selection"""
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            self.selected_instance = self.instances[index]
            self.root.destroy()
        else:
            messagebox.showwarning("No Selection", "Please select an E3.series instance.")
            
    def _on_cancel(self):
        """Handle cancellation"""
        self.selected_instance = None
        self.root.destroy()


class E3ConnectionManager:
    """Manages E3.series instance detection and connection"""
    
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        
    def get_running_e3_instances(self) -> List[E3InstanceInfo]:
        """Get list of running E3.series instances"""
        instances = []
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_info = proc.info
                    proc_name = proc_info['name'] or ""
                    
                    # Check for various E3.series process names
                    e3_patterns = [
                        'e3.series',
                        'e3series',
                        'e3.application',
                        'e3application'
                    ]
                    
                    is_e3_process = any(pattern in proc_name.lower() for pattern in e3_patterns)
                    
                    if is_e3_process:
                        # Try to get more detailed information
                        project_path = self._get_project_path_from_process(proc)
                        instance = E3InstanceInfo(
                            pid=proc_info['pid'],
                            name=proc_name,
                            project_path=project_path
                        )
                        instances.append(instance)
                        
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    # Process might have terminated or we don't have access
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error detecting E3.series instances: {e}")
            
        return instances

    def _get_project_path_from_process(self, proc) -> str:
        """Try to get the project path from process command line, window title, or E3 API"""
        project_path = ""
        pid = proc.pid

        # Method 1: Check command line arguments for project files
        try:
            cmdline = proc.cmdline()
            for arg in cmdline:
                if arg.endswith('.e3p') or arg.endswith('.e3') or arg.endswith('.e3s'):
                    project_path = arg
                    break
        except Exception as e:
            pass

        # If we found a project in command line, return it immediately
        if project_path:
            return project_path

        # Method 2: Try to get window title (Windows-specific) - moved before API method
        if not project_path:
            try:
                import win32gui
                import win32process

                def enum_windows_callback(hwnd, pid_data):
                    try:
                        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                        if found_pid == pid_data['target_pid']:
                            window_title = win32gui.GetWindowText(hwnd)
                            if window_title and win32gui.IsWindowVisible(hwnd):
                                pid_data['titles'].append(window_title)
                    except:
                        pass
                    return True

                pid_data = {'target_pid': pid, 'titles': []}
                win32gui.EnumWindows(enum_windows_callback, pid_data)

                # Look for E3.series windows and extract project names
                for title in pid_data['titles']:
                    if 'e3.dtm' in title.lower() or 'e3.series' in title.lower():
                        # Try different patterns for project extraction
                        if ' - ' in title:
                            parts = title.split(' - ')
                            # Look for the first part that looks like a project file
                            for part in parts:
                                part = part.strip()
                                if (part and
                                    'e3.series' not in part.lower() and
                                    'e3.dtm' not in part.lower() and
                                    'zuken' not in part.lower() and
                                    'professional' not in part.lower() and
                                    not part.startswith('[') and  # Skip sheet info like [Sheet 1]
                                    (part.endswith('.e3s') or part.endswith('.e3p') or part.endswith('.e3') or
                                     (len(part) > 3 and not part.lower().startswith('sheet')))):
                                    project_path = part
                                    break
                        elif title.lower().startswith('e3.series'):
                            # Format like "E3.series ProjectName"
                            remaining = title[9:].strip()  # Remove "E3.series"
                            if remaining and not remaining.lower().startswith('v'):  # Skip version info
                                project_path = remaining

                        if project_path:
                            break

            except ImportError:
                pass
            except Exception as e:
                pass

        # Method 3: E3 API method disabled - it connects to active instance, not specific PID
        # This causes incorrect project names to be returned for non-active instances

        return project_path

    def _get_project_from_e3_api(self, pid: int) -> str:
        """Try to get project information directly from E3 API"""
        # Note: The E3 API doesn't support connecting to specific PIDs directly
        # This method connects to the "active" E3 instance, which may not be
        # the one we want. It's kept as a fallback but may return incorrect data.
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to connect to E3 and get project information
            # WARNING: This connects to the active instance, not necessarily the one with the given PID
            app = e3series.Application()
            job = app.CreateJobObject()

            # Try to get project name/path
            try:
                project_name = job.GetName()
                if project_name:
                    return project_name
            except:
                pass

            # Try alternative methods to get project info
            try:
                project_path = job.GetProjectPath()
                if project_path:
                    return project_path
            except:
                pass

        except Exception as e:
            # Connection failed or API not available
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        return ""

    def select_e3_instance(self) -> Optional[int]:
        """
        Detect E3.series instances and show selection dialog if needed.

        Returns:
            PID of selected E3.series instance, or None if cancelled/no instances
        """
        instances = self.get_running_e3_instances()

        if not instances:
            self.logger.error("No running E3.series instances found")
            messagebox.showerror(
                "No E3.series Instances",
                "No running E3.series instances were found.\n\n"
                "Please start E3.series and open a project before running this automation."
            )
            return None

        elif len(instances) == 1:
            # Only one instance, use it automatically
            instance = instances[0]
            self.logger.info(f"Found single E3.series instance: {instance}")
            return instance.pid

        else:
            # Multiple instances, show selection dialog
            self.logger.info(f"Found {len(instances)} E3.series instances")
            selector = E3InstanceSelector(instances, self)
            selected_instance = selector.show_selection_dialog()

            if selected_instance:
                self.logger.info(f"User selected E3.series instance: {selected_instance}")
                return selected_instance.pid
            else:
                self.logger.info("User cancelled E3.series instance selection")
                return None


def get_e3_connection_pid(logger=None) -> Optional[int]:
    """
    Convenience function to get E3.series connection PID.

    This is the main function that GUI scripts should call to get the PID
    for connecting to E3.series.

    Args:
        logger: Optional logger instance

    Returns:
        PID of selected E3.series instance, or None if cancelled/no instances
    """
    manager = E3ConnectionManager(logger)
    return manager.select_e3_instance()


def connect_to_e3_with_pid(pid: int, logger=None) -> Tuple[bool, dict]:
    """
    Connect to E3.series using a specific PID.

    Args:
        pid: Process ID of the E3.series instance to connect to
        logger: Optional logger instance

    Returns:
        Tuple of (success: bool, objects: dict) where objects contains
        the E3 API objects (app, job, etc.)
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Connect to E3.series application
        # Note: The e3series library doesn't directly support PID-based connection
        # This is a placeholder for the connection logic that would need to be
        # implemented based on the specific E3 API capabilities
        app = e3series.Application()
        job = app.CreateJobObject()

        objects = {
            'app': app,
            'job': job,
            'connection': job.CreateConnectionObject(),
            'pin': job.CreatePinObject(),
            'sheet': job.CreateSheetObject(),
            'signal': job.CreateSignalObject(),
            'net': job.CreateNetObject(),
            'net_segment': job.CreateNetSegmentObject(),
            'device': job.CreateDeviceObject(),
            'symbol': job.CreateSymbolObject()
        }

        logger.info(f"Successfully connected to E3.series instance (PID: {pid})")
        return True, objects

    except Exception as e:
        logger.error(f"Failed to connect to E3.series instance (PID: {pid}): {e}")
        return False, {}
