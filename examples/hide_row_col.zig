//
// An example of how to hide rows and columns using the libxlsxwriter
// library.
//
// In order to hide rows without setting each one, (of approximately 1 million
// rows), Excel uses an optimization to hide all rows that don't have data. In
// Libxlsxwriter we replicate that using the worksheet_set_default_row()
// function.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // Create a new workbook and add a worksheet
    const workbook = xlsxwriter.workbook_new("zig-hide_row_col.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Write some data
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 3, "Some hidden columns.", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 7, 0, "Some hidden rows.", null);

    // Hide all rows without data
    _ = xlsxwriter.worksheet_set_default_row(worksheet, 15, xlsxwriter.LXW_TRUE);

    // Set the height of empty rows that we want to display even if it is
    // the default height
    var row: xlsxwriter.lxw_row_t = 1;
    while (row <= 6) : (row += 1) {
        _ = xlsxwriter.worksheet_set_row(worksheet, row, 15, null);
    }

    // Columns can be hidden explicitly. This doesn't increase the file size
    var options = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 0,
        .collapsed = 0,
    };

    // Use COLS macro equivalent for "G:XFD" range
    _ = xlsxwriter.worksheet_set_column_opt(worksheet, 6, 16383, 8.43, null, &options);

    _ = xlsxwriter.workbook_close(workbook);
}
