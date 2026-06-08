#!/usr/bin/env python3
"""
E3 NA Standards Automation

Runs four E3.series automation scripts with real-time log feedback:
  1. Wire Number Automation
  2. Wire Core Sync
  3. Device Designation Automation
  4. Terminal Pin Name Automation

Author: Jonathan Callahan
Date: 2025-07-11
"""

import logging
import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext

import customtkinter as ctk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from lib.e3_device_designation import run_device_designation_automation
    from lib.e3_terminal_pin_names import run_terminal_pin_name_automation
    from lib.e3_wire_core_sync import WireCoreSynchronizer
    from lib.e3_wire_numbering import run_wire_number_automation
    from lib.e3_field_connection import run_field_connection_automation
    from lib.e3_selector_widget import E3SelectorWidget
except ImportError as exc:
    print(f"Error importing required modules: {exc}")
    sys.exit(1)

try:
    from lib.window_icon import set_window_icon
except ImportError:
    def set_window_icon(root, icon_name):
        return False

try:
    from lib.theme_utils import apply_theme
except ImportError:
    def apply_theme(theme_name="red", appearance_mode="dark"):
        ctk.set_appearance_mode(appearance_mode)
        ctk.set_default_color_theme("blue")


class LogHandler(logging.Handler):
    """Routes log records into a ScrolledText widget on the main thread."""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            msg = self.format(record)
            self.text_widget.after(0, self._append, msg)
        except Exception:
            self.handleError(record)

    def _append(self, msg):
        try:
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.see(tk.END)
        except Exception:
            pass


class E3AutomationGUI(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("E3 NA Standards Automation")
        set_window_icon(self, "e3_na_standards")
        self.geometry("1200x640")
        self.minsize(1060, 520)

        apply_theme("red", "dark")

        self.running_operation = False

        self._build_ui()
        self._setup_logging()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)   # log expands

        # Header: title/description on the left, E3 selector on the right
        header = ctk.CTkFrame(self)
        header.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.grid(row=0, column=0, padx=15, pady=12, sticky="w")

        ctk.CTkLabel(
            left, text="E3 NA Standards Automation", font=("Arial", 22, "bold")
        ).pack(anchor="w")
        ctk.CTkLabel(
            left,
            text="Run E3.series automation scripts with real-time feedback",
            font=("Arial", 13),
        ).pack(anchor="w", pady=(4, 0))

        self._e3_selector = E3SelectorWidget(header)
        self._e3_selector.grid(row=0, column=1, padx=15, pady=12, sticky="e")

        # Buttons
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        btn_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        red_btn = {
            "width": 190, "height": 60,
            "font": ("Arial", 13, "bold"),
            "fg_color": "#C53F3F", "hover_color": "#A02222",
        }

        self.device_designation_btn = ctk.CTkButton(
            btn_frame, text="Set Device\nDesignations",
            command=self.run_device_designation, **red_btn
        )
        self.device_designation_btn.grid(row=0, column=0, padx=8, pady=18)

        self.terminal_pin_btn = ctk.CTkButton(
            btn_frame, text="Set Terminal\nPin Names",
            command=self.run_terminal_pin_names, **red_btn
        )
        self.terminal_pin_btn.grid(row=0, column=1, padx=8, pady=18)

        self.wire_numbers_btn = ctk.CTkButton(
            btn_frame, text="Set Wire\nNumbers",
            command=self.run_wire_numbers, **red_btn
        )
        self.wire_numbers_btn.grid(row=0, column=2, padx=8, pady=18)

        self.wire_core_sync_btn = ctk.CTkButton(
            btn_frame, text="Sync Wire\nCore Properties",
            command=self.run_wire_core_sync, **red_btn
        )
        self.wire_core_sync_btn.grid(row=0, column=3, padx=8, pady=18)

        self.field_connection_btn = ctk.CTkButton(
            btn_frame, text="Set Field\nConnections",
            command=self.run_field_connection, **red_btn
        )
        self.field_connection_btn.grid(row=0, column=4, padx=8, pady=18)

        self.run_all_btn = ctk.CTkButton(
            btn_frame,
            text="Run All\n(Wire → Core → Device → Terminal → Field)",
            command=self.run_all_automation,
            width=210, height=60, font=("Arial", 11, "bold"),
            fg_color="#2AA876", hover_color="#1F8A5F",
        )
        self.run_all_btn.grid(row=0, column=5, padx=8, pady=18)

        # Log area
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_frame, text="Operation Log:", font=("Arial", 15, "bold"), anchor="w"
        ).grid(row=0, column=0, padx=20, pady=(16, 8), sticky="ew")

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, state="disabled",
            bg="#2b2b2b", fg="#ffffff",
            font=("Consolas", 10), insertbackground="#ffffff",
        )
        self.log_text.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="nsew")

        # Status bar
        status_bar = ctk.CTkFrame(self)
        status_bar.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            status_bar, text="Ready — select an operation to begin",
            anchor="w", font=("Arial", 12),
        )
        self.status_label.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        ctk.CTkButton(
            status_bar, text="Clear Log", command=self._clear_log,
            width=100, height=30, fg_color="#666666", hover_color="#555555",
        ).grid(row=0, column=1, padx=20, pady=10)

    def _setup_logging(self):
        self.logger = logging.getLogger("E3AutomationGUI")
        self.logger.setLevel(logging.INFO)
        handler = LogHandler(self.log_text)
        handler.setLevel(logging.INFO)
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        self.logger.addHandler(handler)
        self.logger.info("E3 NA Standards Automation started")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")
        self.logger.info("Log cleared")

    def _set_buttons_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn in (
            self.device_designation_btn, self.terminal_pin_btn,
            self.wire_numbers_btn, self.wire_core_sync_btn,
            self.field_connection_btn, self.run_all_btn,
        ):
            btn.configure(state=state)

    def _set_status(self, message: str, color: str = "#FFFFFF"):
        self.status_label.configure(text=message, text_color=color)

    def _get_pid(self) -> int | None:
        """Return the selected E3 PID, or show an error and return None."""
        pid = self._e3_selector.get_selected_pid()
        if pid is None:
            messagebox.showerror(
                "No Instance",
                "No E3.series instance selected.\n\n"
                "Start E3.series and use the Refresh button if needed.",
            )
        return pid

    def _wire_core_sync_operation(self, logger, pid: int) -> bool:
        """Adapter so WireCoreSynchronizer matches the (logger, pid) -> bool signature."""
        try:
            return WireCoreSynchronizer(e3_pid=pid).run()
        except Exception as exc:
            logger.error(f"Wire core synchronization failed: {exc}")
            return False

    # ------------------------------------------------------------------
    # Generic thread runner
    # ------------------------------------------------------------------

    def _start_operation(self, operation_func, operation_name: str):
        """Get PID on the main thread, then hand off to a daemon worker."""
        if self.running_operation:
            return
        pid = self._get_pid()
        if pid is None:
            return

        self.running_operation = True
        self._set_buttons_enabled(False)
        self._set_status(f"Running {operation_name}...", "#FFA500")
        self.logger.info(f"Starting {operation_name}")

        threading.Thread(
            target=self._run_in_thread,
            args=(operation_func, operation_name, pid),
            daemon=True,
        ).start()

    def _run_in_thread(self, operation_func, operation_name: str, pid: int):
        try:
            success = operation_func(self.logger, pid)
            self.after(0, self._on_done, operation_name, success, None)
        except Exception as exc:
            self.after(0, self._on_done, operation_name, False, str(exc))

    def _on_done(self, operation_name: str, success: bool, error: str | None):
        self.running_operation = False
        self._set_buttons_enabled(True)
        if error:
            self.logger.error(f"Error during {operation_name}: {error}")
            self._set_status(f"Error during {operation_name}", "#FF0000")
            messagebox.showerror("Error", f"Error during {operation_name}:\n{error}")
        elif success:
            self.logger.info(f"{operation_name} completed successfully!")
            self._set_status(f"{operation_name} completed successfully", "#00AA00")
        else:
            self.logger.error(f"{operation_name} failed!")
            self._set_status(f"{operation_name} failed", "#FF0000")
            messagebox.showerror("Error", f"{operation_name} failed! Check the log for details.")

    # ------------------------------------------------------------------
    # Individual operation buttons
    # ------------------------------------------------------------------

    def run_device_designation(self):
        self._start_operation(run_device_designation_automation, "Device Designation Automation")

    def run_terminal_pin_names(self):
        self._start_operation(run_terminal_pin_name_automation, "Terminal Pin Name Automation")

    def run_wire_numbers(self):
        self._start_operation(run_wire_number_automation, "Wire Number Automation")

    def run_wire_core_sync(self):
        self._start_operation(self._wire_core_sync_operation, "Wire Core Synchronization")

    def run_field_connection(self):
        self._start_operation(run_field_connection_automation, "Field Connection Tagging")

    # ------------------------------------------------------------------
    # Run All
    # ------------------------------------------------------------------

    def run_all_automation(self):
        if self.running_operation:
            return
        pid = self._get_pid()
        if pid is None:
            return

        self.running_operation = True
        self._set_buttons_enabled(False)
        self._set_status("Running all automation scripts...", "#FFA500")

        threading.Thread(
            target=self._run_all_in_thread, args=(pid,), daemon=True
        ).start()

    def _run_all_in_thread(self, pid: int):
        operations = [
            (run_wire_number_automation,        "Wire Number Automation"),
            (self._wire_core_sync_operation,    "Wire Core Synchronization"),
            (run_device_designation_automation, "Device Designation Automation"),
            (run_terminal_pin_name_automation,  "Terminal Pin Name Automation"),
            (run_field_connection_automation,   "Field Connection Tagging"),
        ]

        total = len(operations)
        for i, (func, name) in enumerate(operations, 1):
            self.after(0, self._set_status, f"Running {name} ({i}/{total})...", "#FFA500")
            self.logger.info(f"Starting {name} ({i}/{total})")
            try:
                success = func(self.logger, pid)
            except Exception as exc:
                self.after(0, self._on_done, f"All Automation — {name}", False, str(exc))
                return
            if not success:
                self.after(0, self._on_done, f"All Automation — {name}", False, None)
                return

        self.after(0, self._on_done, "All Automation", True, None)


# ---------------------------------------------------------------------------

def main():
    try:
        app = E3AutomationGUI()
        app.mainloop()
    except Exception as exc:
        print(f"Error starting E3 Automation GUI: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
