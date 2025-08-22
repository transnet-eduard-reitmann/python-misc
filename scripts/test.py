#!/usr/bin/env python3
"""
Test script for basic Python functionality testing.

This script demonstrates basic Python features and can be used
to verify the environment is working correctly.
"""

import sys
import datetime
from pathlib import Path


def main():
    """Main function to run basic tests."""
    print("Python Test Script")
    print("=" * 30)
    
    # Python version info
    print(f"Python Version: {sys.version}")
    print(f"Python Executable: {sys.executable}")
    
    # Current working directory
    current_dir = Path.cwd()
    print(f"Current Directory: {current_dir}")
    
    # Timestamp
    now = datetime.datetime.now()
    print(f"Current Time: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Basic calculations
    test_numbers = [1, 2, 3, 4, 5]
    sum_result = sum(test_numbers)
    print(f"Sum of {test_numbers}: {sum_result}")
    
    # File system check
    script_path = Path(__file__)
    print(f"Script Location: {script_path}")
    print(f"Script Size: {script_path.stat().st_size} bytes")
    
    print("\nTest completed successfully! ✓")


if __name__ == "__main__":
    main()