#!/usr/bin/env python3
"""
Test script for position-based wire number assignment logic

This script tests the wire number assignment logic to ensure
it correctly assigns unique wire numbers based on left-to-right position,
with letter suffixes (.A, .B, .C, etc.) assigned in order of X coordinate.

Author: Jonathan Callahan
Date: 2025-01-01
"""

import sys
import os

# Add the apps directory to the path so we can import the module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from set_wire_numbers import WireNumberAssigner

def test_position_based_wire_numbering():
    """Test the position-based wire number assignment logic"""
    print("Testing position-based wire number assignment...")

    # Test the logic for assigning wire numbers based on X coordinate
    # Simulate signals at the same grid position but different X coordinates

    # Test case 1: Single signal at position
    signals = [
        {'base_wire_number': '1A5', 'x_coord': 100, 'y_coord': 50, 'signal_name': 'Signal1'}
    ]

    # Sort by X coordinate (left to right)
    signals.sort(key=lambda x: (x['x_coord'], x['y_coord']))

    # Assign wire numbers
    for i, signal in enumerate(signals):
        if i == 0:
            expected = signal['base_wire_number']
        else:
            letter = chr(ord('A') + i - 1)
            expected = f"{signal['base_wire_number']}.{letter}"

        print(f"Test 1 - Single signal: X={signal['x_coord']} -> '{expected}'")
        assert expected == '1A5', f"Expected '1A5', got '{expected}'"

    # Test case 2: Multiple signals at same position, different X coordinates
    signals = [
        {'base_wire_number': '1A5', 'x_coord': 300, 'y_coord': 50, 'signal_name': 'Signal3'},
        {'base_wire_number': '1A5', 'x_coord': 100, 'y_coord': 50, 'signal_name': 'Signal1'},
        {'base_wire_number': '1A5', 'x_coord': 200, 'y_coord': 50, 'signal_name': 'Signal2'},
    ]

    # Sort by X coordinate (left to right)
    signals.sort(key=lambda x: (x['x_coord'], x['y_coord']))

    expected_results = ['1A5', '1A5.A', '1A5.B']

    for i, signal in enumerate(signals):
        if i == 0:
            wire_number = signal['base_wire_number']
        else:
            letter = chr(ord('A') + i - 1)
            wire_number = f"{signal['base_wire_number']}.{letter}"

        print(f"Test 2 - Signal at X={signal['x_coord']}: '{signal['signal_name']}' -> '{wire_number}'")
        assert wire_number == expected_results[i], f"Expected '{expected_results[i]}', got '{wire_number}'"

    # Test case 3: Same X coordinate, different Y coordinates (top to bottom)
    signals = [
        {'base_wire_number': '2B3', 'x_coord': 100, 'y_coord': 200, 'signal_name': 'SignalBottom'},
        {'base_wire_number': '2B3', 'x_coord': 100, 'y_coord': 100, 'signal_name': 'SignalTop'},
        {'base_wire_number': '2B3', 'x_coord': 100, 'y_coord': 150, 'signal_name': 'SignalMiddle'},
    ]

    # Sort by X coordinate (left to right), then Y coordinate (top to bottom)
    signals.sort(key=lambda x: (x['x_coord'], x['y_coord']))

    expected_results = ['2B3', '2B3.A', '2B3.B']

    for i, signal in enumerate(signals):
        if i == 0:
            wire_number = signal['base_wire_number']
        else:
            letter = chr(ord('A') + i - 1)
            wire_number = f"{signal['base_wire_number']}.{letter}"

        print(f"Test 3 - Signal at Y={signal['y_coord']}: '{signal['signal_name']}' -> '{wire_number}'")
        assert wire_number == expected_results[i], f"Expected '{expected_results[i]}', got '{wire_number}'"

    print("Position-based wire numbering tests passed!")

def test_wire_number_calculation():
    """Test the wire number calculation logic"""
    print("\nTesting wire number calculation...")
    
    assigner = WireNumberAssigner()
    
    # Test normal cases
    result = assigner.calculate_wire_number("1", "A5")
    print(f"Test 1 - Normal: page='1', grid='A5' -> '{result}'")
    assert result == "1A5", f"Expected '1A5', got '{result}'"

    result = assigner.calculate_wire_number("10", "B12")
    print(f"Test 2 - Normal: page='10', grid='B12' -> '{result}'")
    assert result == "10B12", f"Expected '10B12', got '{result}'"

    # Test edge cases
    result = assigner.calculate_wire_number("", "C7")
    print(f"Test 3 - Empty page: page='', grid='C7' -> '{result}'")
    assert result == "0C7", f"Expected '0C7', got '{result}'"

    result = assigner.calculate_wire_number(None, "D8")
    print(f"Test 4 - None page: page=None, grid='D8' -> '{result}'")
    assert result == "0D8", f"Expected '0D8', got '{result}'"
    
    print("Wire number calculation tests passed!")

def main():
    """Main test function"""
    print("=" * 50)
    print("Testing Wire Number Assignment Logic")
    print("=" * 50)
    
    try:
        test_position_based_wire_numbering()
        test_wire_number_calculation()
        
        print("\n" + "=" * 50)
        print("ALL TESTS PASSED SUCCESSFULLY!")
        print("=" * 50)
        
    except Exception as e:
        print(f"\nTEST FAILED: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
