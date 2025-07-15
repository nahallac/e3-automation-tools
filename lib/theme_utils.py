"""
Theme Utilities Module

This module provides utilities for managing themes in the Engineering Tools suite.
"""

import os
import sys
import customtkinter as ctk

def get_theme_path(theme_name):
    """
    Get the absolute path to a theme file.
    
    Args:
        theme_name: The name of the theme file (without .json extension)
        
    Returns:
        str: The absolute path to the theme file
    """
    # Get the application directory
    app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Construct the path to the theme file
    theme_path = os.path.join(app_dir, "resources", "themes", f"{theme_name}.json")
    
    # Check if the theme file exists
    if not os.path.exists(theme_path):
        print(f"Warning: Theme file {theme_path} not found. Using default theme.")
        return None
    
    return theme_path

def apply_theme(theme_name="red", appearance_mode="dark"):
    """
    Apply a theme to the application.
    
    Args:
        theme_name: The name of the theme to apply (default: "red")
        appearance_mode: The appearance mode to use (default: "dark")
    """
    # Set appearance mode
    ctk.set_appearance_mode(appearance_mode)
    
    # Get the theme path
    theme_path = get_theme_path(theme_name)
    
    # If the theme file exists, set it as the default color theme
    if theme_path:
        ctk.set_default_color_theme(theme_path)
    else:
        # Fall back to a built-in theme
        ctk.set_default_color_theme("blue")
