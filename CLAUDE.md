# CLAUDE.md - zig-xlsxwriter Documentation

## Project Overview

**zig-xlsxwriter** is a Zig language binding for the libxlsxwriter C library, enabling programmatic creation of Excel 2007+ XLSX files. This project provides a complete Zig wrapper around libxlsxwriter's functionality with comprehensive examples and verification tools.

**Key Characteristics:**
- Creates Excel files from scratch (cannot modify existing files)
- Full feature coverage: worksheets, formulas, charts, formatting, images, macros
- 68 comprehensive examples demonstrating all major features
- Cross-platform support (Linux, macOS, Windows)
- Licensed under Apache License 2.0

## Architecture

### Core Components

```
┌─────────────────────────────────────────────────────────┐
│                    Zig Application                       │
├─────────────────────────────────────────────────────────┤
│  xlsxwriter module    │  errors module  │  mktmp module │
│  (C FFI bindings)     │  (Error types)  │  (Temp files) │
├─────────────────────────────────────────────────────────┤
│                   libxlsxwriter (C)                      │
├─────────────────────────────────────────────────────────┤
│                  XLSX File Output                        │
└─────────────────────────────────────────────────────────┘
```

### Module Breakdown

#### 1. `src/xlsxwriter.zig` - C Library Bindings
The main wrapper that uses Zig's `@cImport` to expose the libxlsxwriter C API:

```zig
pub const xlsxError = @import("errors.zig");
pub usingnamespace @cImport({
    @cDefine("struct_headname", "");
    @cInclude("xlsxwriter.h");
});
```

This provides transparent access to all libxlsxwriter functions and structures.

#### 2. `src/errors.zig` - Error Handling
Defines 13 Zig-idiomatic error types:

- `RowColumnLimitError` - Row/column exceeds Excel limits
- `SheetnameCannotBeBlank` - Empty worksheet name
- `SheetnameLengthExceeded` - Name too long (>31 chars)
- `MergeRangeOverlaps` - Overlapping merge ranges
- `UnknownUrlType` - Invalid hyperlink URL
- And 8 more...

Includes `formatErr()` function for human-readable error messages.

#### 3. `src/mktmp.zig` - Temporary File Management
Handles embedded resources (images, VBA files) that need file paths:

```zig
pub const TmpFile = struct {
    file: fs.File,
    path: []const u8,
    allocator: mem.Allocator,

    pub fn create(allocator: mem.Allocator, prefix: []const u8) !TmpFile
    pub fn write(self: *TmpFile, data: []const u8) !void
    pub fn readAll(self: *TmpFile, allocator: mem.Allocator) ![]u8
    pub fn cleanUp(self: *TmpFile) void
}
```

**Features:**
- Cross-platform temp directory resolution
- Cryptographically random filenames
- Automatic cleanup with `defer`

## How It Works

### Basic Workflow

1. **Create Workbook:** Initialize a new XLSX file
2. **Add Worksheets:** Create one or more sheets
3. **Write Data:** Add content, formulas, charts, images
4. **Apply Formatting:** Set fonts, colors, borders, alignment
5. **Close Workbook:** Finalize and save the file

### Code Pattern Example

```zig
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // 1. Create workbook
    const workbook = xlsxwriter.workbook_new("output.xlsx");
    defer _ = xlsxwriter.workbook_close(workbook);

    // 2. Add worksheet
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // 3. Write data
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Hello", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 123.45, null);

    // 4. Add formula
    _ = xlsxwriter.worksheet_write_formula(worksheet, 2, 0, "=A2*2", null);

    // 5. Apply formatting
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 0, "Bold Text", bold);
}
```

### Advanced Patterns

#### Embedded Images

Images are embedded at compile time and extracted to temporary files:

```zig
const mktmp = @import("mktmp");

const logo_data = @embedFile("logo.png");

var gpa = std.heap.GeneralPurposeAllocator(.{}){};
const allocator = gpa.allocator();

var tmp_file = try mktmp.TmpFile.create(allocator, "logo_");
defer tmp_file.cleanUp();

try tmp_file.write(logo_data);
_ = xlsxwriter.worksheet_insert_image(worksheet, row, col, tmp_file.path.ptr);
```

#### Chart Creation

```zig
// Create chart
var chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_LINE);

// Add data series
var series = xlsxwriter.chart_add_series(
    chart,
    "=Sheet1!$A$2:$A$7",  // Categories
    "=Sheet1!$B$2:$B$7"   // Values
);

// Configure series
_ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

// Insert chart into worksheet
_ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 4, chart);
```

#### VBA Macros

```zig
const vba_data = @embedFile("vbaProject.bin");

var tmp_file = try mktmp.TmpFile.create(allocator, "vba_");
defer tmp_file.cleanUp();

try tmp_file.write(vba_data);
const c_path = @as([*c]const u8, @ptrCast(tmp_file.path.ptr));

// Must use .xlsm extension for macro-enabled files
const workbook = xlsxwriter.workbook_new("output.xlsm");
_ = xlsxwriter.workbook_add_vba_project(workbook, c_path);
```

#### Conditional Formatting

```zig
var conditional_format: xlsxwriter.lxw_conditional_format =
    std.mem.zeroes(xlsxwriter.lxw_conditional_format);

conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_LESS_THAN;
conditional_format.value = 33;

const format = xlsxwriter.workbook_add_format(workbook);
_ = xlsxwriter.format_set_bg_color(format, xlsxwriter.LXW_COLOR_RED);

conditional_format.format = format;

_ = xlsxwriter.worksheet_conditional_format_range(
    worksheet,
    first_row, first_col,
    last_row, last_col,
    &conditional_format
);
```

## Build System

### Build Configuration (`build.zig`)

The build system automatically discovers and compiles all 68 examples:

```zig
// Discovers all .zig files in examples/
// Creates build targets for each
// Provides run targets with -Dexample=name
```

### Common Build Commands

```bash
# Build all examples
zig build

# Build and run specific example
zig build run -Dexample=hello

# Build specific example only
zig build hello

# Clean build artifacts
zig build clean

# Verbose build with summary
zig build --summary all -freference-trace
```

### Dependencies (`build.zig.zon`)

```zig
.dependencies = .{
    .xlsxwriter = .{
        .url = "git+https://github.com/jmcnamara/libxlsxwriter#caf4158...",
        .hash = "122080b8864c9fcc...",
    },
}
```

## Examples

The project includes **68 examples** covering all major features:

### Basic Examples
- `hello.zig` - Minimal 16-line example
- `tutorial1.zig` - Basic data and formulas
- `tutorial2.zig` - Adding formats
- `tutorial3.zig` - Complex operations

### Data Writing
- `dates_and_times01-04.zig` - Date/time handling
- `utf8.zig` - UTF-8 text support
- `dynamic_arrays.zig` - Variable-length data
- `hyperlinks.zig` - URL linking

### Formatting
- `format_font.zig` - Font properties
- `format_num_format.zig` - Number formatting
- `background.zig` - Cell backgrounds
- `diagonal_border.zig` - Border styling
- `rich_strings.zig` - Mixed format text

### Charts (12 Types)
- `chart_line.zig`, `chart_area.zig`, `chart_column.zig`
- `chart_bar.zig`, `chart_pie.zig`, `chart_doughnut.zig`
- `chart_scatter.zig`, `chart_radar.zig`, `chart_stock.zig`
- `chart_clustered.zig`, `chart_data_table.zig`, `chart_pattern.zig`

### Advanced Features
- `images.zig` - Image embedding with positioning
- `macro.zig` - VBA macro integration
- `comments1-2.zig` - Cell comments
- `conditional_format1-2.zig` - Conditional formatting
- `data_validate.zig` - Data validation dropdowns
- `merge_range.zig` - Merged cells
- `tables.zig` - Excel table creation
- `worksheet_protection.zig` - Sheet protection
- `outline_*.zig` - Row/column grouping
- `panes.zig` - Frozen panes
- `headers_footers.zig` - Page setup
- `watermark.zig` - Background watermarks

## Verification System

### Example Status Tracking

```bash
# Check all examples
python3 utils/evaluate.py

# Check specific example
python3 utils/evaluate.py example_name
```

**Status Levels:**
- ✅ DONE - Fully implemented and visually verified
- ⚠️  IN PROGRESS - Implemented but not verified
- ❌ NOT STARTED - Not yet implemented

**Exit Codes:**
- `0` - Fully verified
- `1` - Needs verification
- `2` - Not implemented

### Visual Verification

```bash
# Create screenshots for comparison
python3 utils/create_screenshots.py example_name
```

Compares Zig output with C implementation output using PIL for image processing.

### Example Converter

```bash
# Convert C example to Zig skeleton
zig build-exe utils/convert_example.zig
./convert_example examples/c/source.c examples/target.zig
```

## Key Constants and Enums

### Colors
```zig
LXW_COLOR_BLACK, LXW_COLOR_RED, LXW_COLOR_BLUE, LXW_COLOR_GREEN, etc.
```

### Chart Types
```zig
LXW_CHART_LINE, LXW_CHART_AREA, LXW_CHART_BAR, LXW_CHART_PIE,
LXW_CHART_DOUGHNUT, LXW_CHART_SCATTER, LXW_CHART_RADAR, etc.
```

### Alignment
```zig
LXW_ALIGN_CENTER, LXW_ALIGN_LEFT, LXW_ALIGN_RIGHT,
LXW_ALIGN_VERTICAL_CENTER, LXW_ALIGN_VERTICAL_TOP, etc.
```

### Borders
```zig
LXW_BORDER_THIN, LXW_BORDER_MEDIUM, LXW_BORDER_THICK,
LXW_BORDER_DOUBLE, LXW_BORDER_DASHED, etc.
```

## Feature Capabilities

### Supported Operations
- ✅ Multiple worksheets per workbook
- ✅ Formulas and array formulas
- ✅ 12 chart types with styling
- ✅ Image embedding (PNG, JPG, BMP, GIF)
- ✅ VBA macro integration
- ✅ Data validation with dropdowns
- ✅ Conditional formatting rules
- ✅ Table creation with filtering
- ✅ Hyperlinks and external references
- ✅ Worksheet protection
- ✅ Headers and footers
- ✅ Merged cells
- ✅ Comments and notes
- ✅ Outline/grouping
- ✅ Frozen panes
- ✅ Rich text formatting
- ✅ Watermarks
- ✅ Named ranges
- ✅ Custom properties and metadata

### Performance
- Constant memory mode for large datasets
- Streaming writes to file
- Large file support (10,000+ rows tested)

## Development Workflow

### Working on Examples

1. **Identify next example:**
   ```bash
   python3 utils/evaluate.py
   ```

2. **Implement the example:**
   - Reference C version in `examples/c/{name}.c`
   - Zig skeleton in `examples/{name}.zig`
   - Output file: `zig-{name}.xlsx`

3. **Build and run:**
   ```bash
   zig build run -Dexample=example_name
   ```

4. **Verify output:**
   ```bash
   python3 utils/create_screenshots.py example_name
   python3 utils/evaluate.py example_name
   ```

### Coding Guidelines

- Use `[:0]const u8` for null-terminated strings (preferred over `[*c]const u8`)
- Embed resources with `@embedFile` and extract to temp files
- Output filenames use `zig-{example_name}.xlsx` format
- Clean up temporary files with `defer tmp_file.cleanUp()`
- Follow error handling patterns from `src/errors.zig`

## File Organization

```
zig-xlsxwriter/
├── src/                      # Core Zig source
│   ├── xlsxwriter.zig       # C bindings
│   ├── errors.zig           # Error types
│   └── mktmp.zig            # Temp file handling
├── examples/                 # 68 example programs
│   ├── *.zig                # Zig examples
│   ├── c/                   # C reference examples
│   └── resources/           # Images, VBA files
├── utils/                    # Verification tools
│   ├── evaluate.py          # Status checker
│   ├── create_screenshots.py # Visual verification
│   └── convert_example.zig  # C→Zig converter
├── testing/                  # Test artifacts
│   ├── screenshots/         # Visual comparisons
│   └── comparison_results/  # Test logs
├── build.zig                # Build configuration
├── build.zig.zon           # Dependencies
└── README.md               # Main documentation
```

## Technical Details

### Memory Management
- Uses Zig's allocator pattern
- C library memory managed by libxlsxwriter
- Defer cleanup for temporary files

### C FFI
- `@cImport` for transparent C interop
- Direct access to C structures
- Pointer conversions with `@ptrCast`
- C strings as `[*:0]const u8`

### File I/O
- Direct workbook creation via libxlsxwriter
- Embedded resource extraction
- Temporary file isolation for path-required resources

## CI/CD

GitHub Actions workflow (`.github/workflows/CI.yml`):
- Multi-platform testing (Ubuntu, macOS, Windows)
- Automated builds on every push
- Build summary reporting
- Early failure detection

## Requirements

### Build Dependencies
- Zig v0.14.0+
- libxlsxwriter (fetched automatically)
- C compiler (for libxlsxwriter)

### Optional (Verification)
- Python 3.6+
- PIL/Pillow (image processing)
- NumPy (numerical operations)
- Microsoft Excel (visual verification)

## Quick Start

```bash
# Clone repository
git clone https://github.com/awesomo4000/zig-xlsxwriter.git
cd zig-xlsxwriter

# Build all examples
zig build

# Run hello world example
zig build run -Dexample=hello

# Output: zig-hello.xlsx
```

## Summary

zig-xlsxwriter is a professional-grade binding that provides:

- ✅ Complete API coverage of libxlsxwriter
- ✅ 68 production examples
- ✅ Comprehensive error handling
- ✅ Cross-platform support
- ✅ Automated verification
- ✅ Clear development workflow
- ✅ Focus on functionality and performance

The codebase demonstrates excellent practices for C FFI in Zig, proper memory management with embedded resources, and comprehensive example coverage for a binding library.
