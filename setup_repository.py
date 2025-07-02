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

def copy_file_if_exists(source, destination):
    """Copy a file if it exists, create destination directory if needed"""
    source_path = Path(source)
    dest_path = Path(destination)
    
    if source_path.exists():
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, dest_path)
        print(f"✓ Copied {source} → {destination}")
        return True
    else:
        print(f"✗ Source file not found: {source}")
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
        "docs", 
        "gui",
        ".github/workflows"
    ]
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
        print(f"✓ Created directory: {directory}")
    
    # Copy files from main repository
    files_to_copy = [
        ("apps/set_wire_numbers.py", "scripts/set_wire_numbers.py"),
        ("apps/test_unique_wire_numbers.py", "scripts/test_unique_wire_numbers.py"),
        ("README_wire_numbering.md", "docs/wire_numbering.md"),
    ]
    
    print("\nCopying files from main repository...")
    for source, dest in files_to_copy:
        source_full = main_repo / source
        copy_file_if_exists(source_full, dest)
    
    # Create a minimal requirements.txt
    print("\nCreating requirements.txt...")
    with open("requirements.txt", "w") as f:
        f.write("# E3.series Automation Tools Dependencies\n")
        f.write("pywin32==306\n")
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
