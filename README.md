# zig-xlsxwriter

The `zig-xlsxwriter` is a wrapper of [`libxlsxwriter`](https://github.com/jmcnamara/libxlsxwriter) that can be used to write text, numbers,
dates and formulas to multiple worksheets in a new Excel 2007+ xlsx file. It
has a focus on performance and on fidelity with the file format created by
Excel. It cannot be used to modify an existing file.


## Requirements

- [zig v0.14.0 or higher](https://ziglang.org/download)
- [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)
- Python 3.6+ with PIL (Pillow) and numpy for verification tools
- Microsoft Excel for visual verification


## Development Workflow

### Building Examples

To build and run a specific example:

```bash
zig build run -Dexample=hello
```

### Implementing the Zig versions of the C examples:

- The examples/c/*.c files are the C language examples that are being converted
  and verified. They are good to read when determining how to implement a new
  example.
- The examples/*.zig files are either completed .zig versions or skeletons left
  to implement.
- The skeleton files have the proper output filenames for the Zig versions.
  They are all formatted as zig-{example}.xlsx .
- The source code for libxlsxwriter is in the $HOME/src/libxlsxwriter/ directory.
- Try to avoid using [*c] , for null terminated C strings, prefer [:0]const u8 instead of [*c]const u8
- For examples that read a file at runtime, to test the functionality, we're using zig's @embedFile to put the file in the binary, then extracting it to a temp directory which gets cleaned up. This way runtime testing of examples won't fail if they are run outside of the directory where the original test file is located. 

### Verifying Examples

The project includes a verification workflow to ensure Zig examples match their C counterparts:

1. Check example status and find next one to work on:
   ```bash
   # List all examples and their status
   python3 utils/evaluate.py
   
   # Or check a specific example
   python3 utils/evaluate.py example_name
   ```

2. After implementing/modifying an example, verify it:
   ```bash
   # Build and run the example
   zig build run -Dexample=example_name

   # Create screenshots and verify visual match
   python3 utils/create_screenshots.py example_name

   # Confirm full verification
   python3 utils/evaluate.py example_name
   ```

**Important Notes**: When porting C examples to Zig, the skeleton files already have the correct output filename. All Zig example output files should start with `zig-` prefix (e.g. `zig-chart_line.xlsx`). Do not modify this prefix when implementing the examples.

The verification process checks:
- If the Zig implementation exists
- If screenshots match between C and Zig outputs
- If the implementation is up to date

Status indicators:
- ✅ DONE: Fully implemented and verified
- ⚠️ IN PROGRESS: Implemented but needs verification
- ❌ NOT STARTED: Not implemented yet

Exit codes from evaluate.py:
- 0: Example is fully implemented and verified
- 1: Example is implemented but needs verification
- 2: Example is not implemented
