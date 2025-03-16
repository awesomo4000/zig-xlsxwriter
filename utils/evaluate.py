#!/usr/bin/env python3
"""
Evaluate whether a solution is done when given a specific example.
This script determines if an example is fully implemented and verified
by checking if the Zig file exists and if visual outputs match between
C and Zig implementations.
"""

import os
import sys
import datetime
import re
import argparse
import time
from pathlib import Path
import shutil

# Add function to clear terminal screen
def clear_screen():
    """Clear the terminal screen."""
    # Check if we're on Windows or Unix-like
    if os.name == 'nt':  # Windows
        os.system('cls')
    else:  # Unix-like
        os.system('clear')


def get_example_status(example_name):
    """
    Determine the status of an example based on implementation and verification.
    
    Returns:
        tuple: (status, message)
            status: "DONE" if fully implemented and verified, "IN PROGRESS" if implemented but not verified, "NOT_STARTED" if not implemented
            message: Detailed message about the status
    """
    # Check if the example file exists
    file_exists, _ = check_example_file_exists(example_name)
    
    # Check if screenshots exist
    screenshots_exist = check_screenshots_exist(example_name)
    
    # Check comparison results
    comparison_match = False
    if screenshots_exist:
        comparison_match, _ = check_comparison_results(example_name)
    
    # Determine status
    if file_exists and screenshots_exist and comparison_match:
        status = "DONE"
        message = f"✅ Example '{example_name}' is fully implemented and verified."
    elif file_exists:
        status = "IN PROGRESS"
        message = f"⚠️ Example '{example_name}' is implemented but not fully verified."
    else:
        status = "NOT_STARTED"
        message = f"❌ Example '{example_name}' is not implemented."
    
    return status, message


def check_example_file_exists(example_name):
    """Check if the example file exists in the examples directory."""
    root_dir = Path(__file__).parent.parent
    example_file = root_dir / "examples" / f"{example_name}.zig"
    
    if example_file.exists():
        return True, f"📄 Example file '{example_name}.zig' exists."
    else:
        return False, f"❌ Example file '{example_name}.zig' does not exist."


def check_c_example_exists(example_name):
    """Check if a corresponding C example exists."""
    root_dir = Path(__file__).parent.parent
    c_example_file = root_dir / "examples" / "c" / f"{example_name}.c"
    
    if c_example_file.exists():
        return True, f"📄 C example file '{example_name}.c' exists."
    else:
        return False, f"❓ C example file '{example_name}.c' not found."


def check_implementation_freshness(example_name):
    """Check if Zig implementation is up-to-date with C implementation."""
    root_dir = Path(__file__).parent.parent
    zig_file = root_dir / "examples" / f"{example_name}.zig"
    c_file = root_dir / "examples" / "c" / f"{example_name}.c"
    
    if not zig_file.exists() or not c_file.exists():
        return None, "Cannot compare timestamps - one or both files missing."
    
    zig_mtime = datetime.datetime.fromtimestamp(zig_file.stat().st_mtime)
    c_mtime = datetime.datetime.fromtimestamp(c_file.stat().st_mtime)
    
    time_diff = zig_mtime - c_mtime
    
    if zig_mtime > c_mtime:
        return True, f"✅ Zig implementation is newer ({abs(time_diff.days)} days, {abs(time_diff.seconds//3600)} hours)"
    else:
        return False, f"⚠️ C implementation is newer ({abs(time_diff.days)} days, {abs(time_diff.seconds//3600)} hours)"


def check_screenshots_exist(example_name):
    """Check if a screenshot exists for the example."""
    root_dir = Path(__file__).parent.parent
    screenshots_dir = root_dir / "testing" / "screenshots"
    
    # Handle special case for conditional_format1
    screenshot_name = example_name
    if example_name == "conditional_format1":
        screenshot_name = "conditional_format_simple"
    
    # Check for combined screenshot
    screenshot_file = screenshots_dir / f"comparison_{screenshot_name}.png"
    
    return screenshot_file.exists()


def check_comparison_results(example_name):
    """Check if comparison results exist and indicate a match."""
    root_dir = Path(__file__).parent.parent
    comparison_dir = root_dir / "testing" / "comparison_results"
    
    # Handle special case for conditional_format1
    result_name = example_name
    if example_name == "conditional_format1":
        result_name = "conditional_format_simple"
    
    result_file = comparison_dir / f"{result_name}_output.txt"
    
    if not result_file.exists():
        return False, f"❌ No comparison results found."
    
    # Read the comparison result file
    with open(result_file, 'r') as f:
        content = f.read()
    
    # Check if the comparison indicates a match
    if "MATCH" in content.upper() or "IDENTICAL" in content.upper() or "SUCCESS" in content.upper():
        return True, f"✅ Visual comparison indicates a match."
    else:
        return False, f"❌ Visual comparison indicates differences."


def compare_screenshots(example_name):
    """Compare screenshots of C and Zig implementations using image similarity."""
    try:
        from PIL import Image
        import numpy as np
    except ImportError:
        return None, "⚠️ PIL or numpy not installed. Cannot perform image comparison."
    
    root_dir = Path(__file__).parent.parent
    screenshots_dir = root_dir / "testing" / "screenshots"
    
    # Handle special case for conditional_format1
    screenshot_name = example_name
    if example_name == "conditional_format1":
        screenshot_name = "conditional_format_simple"
    
    c_screenshot = screenshots_dir / f"c_{screenshot_name}.png"
    zig_screenshot = screenshots_dir / f"zig_{screenshot_name}.png"
    
    if not c_screenshot.exists() or not zig_screenshot.exists():
        return None, "⚠️ Cannot compare - one or both screenshots missing."
    
    try:
        # Load images
        c_img = Image.open(c_screenshot)
        zig_img = Image.open(zig_screenshot)
        
        # Resize to same dimensions for comparison
        if c_img.size != zig_img.size:
            zig_img = zig_img.resize(c_img.size)
        
        # Convert to numpy arrays
        c_array = np.array(c_img)
        zig_array = np.array(zig_img)
        
        # Calculate mean squared error
        mse = np.mean((c_array - zig_array) ** 2)
        
        # Calculate structural similarity (simplified)
        similarity = 1 - (mse / 255**2)
        
        if similarity > 0.95:
            return True, f"✅ Screenshots are visually similar ({similarity:.2%} match)."
        else:
            return False, f"❌ Screenshots differ significantly ({similarity:.2%} match)."
    
    except Exception as e:
        return None, f"⚠️ Error comparing images: {str(e)}"


def is_example_fully_implemented(example_name):
    """
    Check if an example is fully implemented and verified.
    
    Returns:
        tuple: (is_implemented, is_verified, message)
            is_implemented: True if the Zig file exists
            is_verified: True if screenshots exist and match
            message: Detailed message about the status
    """
    # Check if the example file exists
    file_exists, _ = check_example_file_exists(example_name)
    
    # Check if screenshots exist
    screenshots_exist = check_screenshots_exist(example_name)
    
    # Check comparison results
    comparison_match = False
    if screenshots_exist:
        comparison_match, _ = check_comparison_results(example_name)
    
    # Determine implementation status
    is_implemented = file_exists
    is_verified = screenshots_exist and comparison_match
    
    # Generate message
    if is_implemented and is_verified:
        message = f"✅ Example '{example_name}' is fully implemented and verified."
    elif is_implemented and not is_verified:
        message = f"⚠️ Example '{example_name}' is implemented but not fully verified."
    else:
        message = f"❌ Example '{example_name}' is not implemented."
    
    return is_implemented, is_verified, message


def get_all_examples():
    """Get all examples from the examples/c directory."""
    root_dir = Path(__file__).parent.parent
    c_examples_dir = root_dir / "examples" / "c"
    
    # Get all .c files in the examples/c directory
    c_files = c_examples_dir.glob("*.c")
    
    # Extract example names (remove .c extension)
    examples = {file.stem for file in c_files}
    
    return examples


def list_all_examples():
    """List all examples and their status."""
    root_dir = Path(__file__).parent.parent
    
    # Get all examples from the examples/c directory
    all_examples = get_all_examples()
    
    print(f"\n{'=' * 70}")
    print(f"{'EXAMPLE':<30} {'STATUS':<20} {'ZIG':<5} {'SCRN':<5} {'MATCH':<5}")
    print(f"{'-' * 70}")
    
    done_count = 0
    in_progress_count = 0
    not_started_count = 0
    
    for example in sorted(all_examples):
        # Get actual status based on implementation
        actual_status, _ = get_example_status(example)
        
        # Check if Zig file exists
        zig_file = root_dir / "examples" / f"{example}.zig"
        zig_exists = "✅" if zig_file.exists() else "❌"
        
        # Check if screenshots exist
        screenshots_exist = check_screenshots_exist(example)
        screenshots_status = "✅" if screenshots_exist else "❌"
        
        # Check visual match
        if screenshots_exist:
            visual_match, _ = check_comparison_results(example)
            visual_status = "✅" if visual_match else "❌"
        else:
            visual_status = "❓"
        
        # Format status with emoji
        if actual_status == "DONE":
            formatted_status = "DONE"
            done_count += 1
        elif actual_status == "IN PROGRESS":
            formatted_status = "IN PROGRESS"
            in_progress_count += 1
        else:
            formatted_status = "NOT_STARTED"
            not_started_count += 1
        
        print(f"{example:<30} {formatted_status:<20} {zig_exists:<5} {screenshots_status:<5} {visual_status:<5}")
    
    print(f"{'-' * 70}")
    total = len(all_examples)
    print(f"Total: {total} examples ({done_count} done, {in_progress_count} in progress, {not_started_count} not started)")
    print(f"Progress: {done_count/total*100:.1f}% complete,") 
    print(f"{(done_count+in_progress_count)/total*100:.1f}% in progress or complete")
    print(f"{'=' * 70}\n")


def get_c_excel_file(example_name):
    """Get the path to the C Excel file."""
    root_dir = Path(__file__).parent.parent
    c_output_dir = root_dir / "testing" / "c-output-xls"
    
    # Handle special case for conditional_format1
    c_file_name = example_name
    if example_name == "conditional_format1":
        c_file_name = "conditional_format_simple"
    
    # Handle special case for dates_and_times examples
    if example_name.startswith("dates_and_times"):
        c_file_name = example_name.replace("dates_and_times", "date_and_times")
    
    # Special case for macro example which uses .xlsm
    extension = ".xlsm" if example_name == "macro" else ".xlsx"
    
    c_excel_file = c_output_dir / f"{c_file_name}{extension}"
    
    if not c_excel_file.exists():
        print(f"C Excel file not found: {c_excel_file}")
        return None
    
    return c_excel_file


def cleanup_excel_file(example_name, zig_excel_file):
    """
    Move the Zig-generated Excel file to the zig-output-xls directory.
    
    Args:
        example_name: The name of the example
        zig_excel_file: Path to the Zig-generated Excel file
    
    Returns:
        bool: True if the file was moved successfully, False otherwise
    """
    root_dir = Path(__file__).parent.parent
    zig_output_dir = root_dir / "testing" / "zig-output-xls"
    
    # Create the directory if it doesn't exist
    zig_output_dir.mkdir(parents=True, exist_ok=True)
    
    # Special case for macro example which uses .xlsm
    extension = ".xlsm" if example_name == "macro" else ".xlsx"
    
    # Destination path 
    dest_file = zig_output_dir / f"{example_name}{extension}"
    
    try:
        # Move the file
        shutil.move(str(zig_excel_file), str(dest_file))
        print(f"✅ Moved Excel file to: {dest_file}")
        return True
    except Exception as e:
        print(f"❌ Error moving Excel file: {e}")
        return False


def main():
    """Main function to evaluate an example."""
    parser = argparse.ArgumentParser(description="Evaluate the implementation status of examples")
    parser.add_argument("example", nargs="?", help="Example name to evaluate")
    parser.add_argument("--monitor", type=int, nargs="?", const=5, 
                        help="Monitor mode: continuously update status every N seconds (default: 5)")
    parser.add_argument("--cleanup", action="store_true",
                        help="Move generated Excel file to zig-output-xls directory")
    
    args = parser.parse_args()
    
    # Monitor mode
    if args.monitor is not None:
        try:
            # Track previous status of each example
            prev_status = {}
            while True:
                all_examples = get_all_examples()
                current_status = {}
                
                # Get current status for all examples
                for example in all_examples:
                    status, _ = get_example_status(example)
                    current_status[example] = status
                
                # Check if any status changed
                status_changed = False
                for example, status in current_status.items():
                    if example not in prev_status or prev_status[example] != status:
                        status_changed = True
                        break
                
                # Only update display if status changed
                if status_changed:
                    clear_screen()
                    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    print(f"Monitoring examples status (updates on status change)")
                    print(f"Last update: {current_time}")
                    print(f"Press Ctrl+C to exit\n")
                    list_all_examples()
                    prev_status = current_status.copy()
                
                time.sleep(args.monitor)
        except KeyboardInterrupt:
            print("\nMonitoring stopped.")
            return 0
    
    if not args.example:
        # No example specified, list all examples
        list_all_examples()
        return 0

    # Handle cleanup if requested
    if args.cleanup:
        root_dir = Path(__file__).parent.parent
        extension = ".xlsm" if args.example == "macro" else ".xlsx"
        zig_excel_file = root_dir / f"zig-{args.example}{extension}"
        
        if not zig_excel_file.exists():
            print(f"❌ No Excel file found at: {zig_excel_file}")
            return 1
            
        if cleanup_excel_file(args.example, zig_excel_file):
            return 0
        return 1

    # Determine status based on implementation and verification
    status, message = get_example_status(args.example)
    print(message)
    
    # Check if the example file exists
    file_exists, file_message = check_example_file_exists(args.example)
    print(file_message)
    
    # Check if C example exists
    c_exists, c_message = check_c_example_exists(args.example)
    print(c_message)
    
    # Check implementation freshness
    if file_exists and c_exists:
        is_fresh, fresh_message = check_implementation_freshness(args.example)
        print(fresh_message)
    
    # Check if screenshots exist
    screenshots_exist = check_screenshots_exist(args.example)
    screenshots_message = "✅ Screenshots exist" if screenshots_exist else "❌ Screenshots do not exist"
    print(screenshots_message)
    
    # Check comparison results
    if screenshots_exist:
        comparison_match, comparison_message = check_comparison_results(args.example)
        print(comparison_message)
        
        # Perform direct image comparison
        try:
            image_match, image_message = compare_screenshots(args.example)
            if image_match is not None:
                print(image_message)
        except ImportError:
            print("⚠️ PIL or numpy not installed. Skipping direct image comparison.")
    
    # Overall status
    if status == "DONE":
        print("\n🎉 Example is fully implemented and verified!")
        return 0
    elif status == "IN PROGRESS":
        print("\n🔧 Example is implemented but not fully verified.")
        return 1
    else:
        print("\n❌ Example is not implemented.")
        return 2


if __name__ == "__main__":
    sys.exit(main())
