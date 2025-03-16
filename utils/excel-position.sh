#!/bin/bash

# Check if the correct number of arguments is provided
if [ "$#" -ne 2 ]; then
    echo "Usage: $0 <excel_file_path_1> <excel_file_path_2>"
    exit 1
fi

# Get absolute paths for both files
FILE_PATH_1=$(realpath "$1")
FILE_PATH_2=$(realpath "$2")

# Verify that both files exist
if [ ! -f "$FILE_PATH_1" ]; then
    echo "Error: File not found: $FILE_PATH_1"
    exit 1
fi

if [ ! -f "$FILE_PATH_2" ]; then
    echo "Error: File not found: $FILE_PATH_2"
    exit 1
fi

# Use osascript to execute AppleScript that positions Excel windows as requested
osascript <<EOF
set displaySize to do shell script "system_profiler SPDisplaysDataType | grep Resolution | head -1"
set screenWidth to word 2 of displaySize
set screenHeight to word 4 of displaySize
set gap to 0
set halfGap to gap / 2
tell application "Microsoft Excel"
    # Activate Excel (bring to front)
    activate
    
    # Get the screen dimensions to calculate window sizes
    
    # Calculate window dimensions - now 1/5 of screen width
    set windowWidth to screenWidth / 2
    set windowHeight to screenHeight * 7 / 8

    
    # Open the first file and position it at the top left
    open POSIX file "$FILE_PATH_1"
    set bounds of window 1 to {0, 0, windowWidth-halfGap, windowHeight}
    
    # Open the second file and position it
    # We use "window 1" again here because the newly opened window
    # is not referenced by "window 1"
    open POSIX file "$FILE_PATH_2"
    set bounds of window 1 to {windowWidth + halfGap, 0, windowWidth * 2 + halfGap, windowHeight}
end tell
EOF
