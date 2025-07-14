#!/usr/bin/env python3
"""
Setup script for E3 Automation Tools Repository

This script helps set up the initial repository structure and copies
the necessary files from the main work-scripts repository.
"""

import os
import shutil
import sys
from pathlib import Path

def copy_file_if_exists(source, destination, optional=False):
    """Copy a file if it exists, create destination directory if needed"""
    source_path = Path(source)
    dest_path = Path(destination)

    if source_path.exists():
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, dest_path)
        print(f"✓ Copied {source} → {destination}")
        return True
    else:
        if optional:
            print(f"⚠ Optional file not found (skipping): {source}")
        else:
            print(f"✗ Required file not found: {source}")
        return False

def main():
    """Main setup function"""
    print("E3 Automation Tools Repository Setup")
    print("=" * 40)
    
    # Get the path to the main work-scripts repository
    main_repo_path = input("Enter the path to your work-scripts repository: ").strip()
    if not main_repo_path:
        print("Error: Repository path is required")
        sys.exit(1)
    
    main_repo = Path(main_repo_path)
    if not main_repo.exists():
        print(f"Error: Repository path does not exist: {main_repo_path}")
        sys.exit(1)
    
    # Create directory structure
    directories = [
        "scripts",
        "lib",
        "docs",
        "gui",
        ".github/workflows"
    ]
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
        print(f"✓ Created directory: {directory}")
    
    # Copy files from main repository
    files_to_copy = [
        # Legacy individual scripts (optional - may not exist)
        ("apps/set_wire_numbers.py", "scripts/set_wire_numbers.py", True),
        ("apps/test_unique_wire_numbers.py", "scripts/test_unique_wire_numbers.py", True),

        # New library modules (required)
        ("lib/e3_wire_numbering.py", "lib/e3_wire_numbering.py", False),
        ("lib/e3_device_designation.py", "lib/e3_device_designation.py", False),
        ("lib/e3_terminal_pin_names.py", "lib/e3_terminal_pin_names.py", False),

        # NA Standards GUI (required)
        ("apps/e3_NA_Standards.py", "gui/e3_NA_Standards.py", False),

        # Documentation (required)
        ("README_wire_numbering.md", "docs/wire_numbering.md", False),
        ("README_device_designation.md", "docs/device_designation.md", False),
        ("README_terminal_pin_names.md", "docs/terminal_pin_names.md", False),
    ]
    
    print("\nCopying files from main repository...")
    for file_info in files_to_copy:
        if len(file_info) == 3:
            source, dest, optional = file_info
        else:
            source, dest = file_info
            optional = False
        source_full = main_repo / source
        copy_file_if_exists(source_full, dest, optional)
    
    # Create a minimal requirements.txt
    print("\nCreating requirements.txt...")
    with open("requirements.txt", "w") as f:
        f.write("# E3.series Automation Tools Dependencies\n")
        f.write("\n")
        f.write("# Core Windows COM interface for E3.series\n")
        f.write("pywin32==306\n")
        f.write("\n")
        f.write("# E3.series Python API\n")
        f.write("e3series\n")
        f.write("\n")
        f.write("# GUI framework for NA Standards app\n")
        f.write("customtkinter==5.2.1\n")
    print("✓ Created requirements.txt")
    
    print("\n" + "=" * 40)
    print("Setup complete!")
    print("\nNext steps:")
    print("1. Review the copied files")
    print("2. Create a new GitHub repository")
    print("3. Push this content to the new repository")
    print("4. Set up the sync workflows if desired")

if __name__ == "__main__":
    main()
