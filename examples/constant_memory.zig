//
// Example of using libxlsxwriter for writing large files in constant memory
// mode.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const row_max: u32 = 1000;
    const col_max: u16 = 50;

    // Set the workbook options
    var options = xlsxwriter.lxw_workbook_options{
        .constant_memory = xlsxwriter.LXW_TRUE,
        .tmpdir = null,
        .use_zip64 = xlsxwriter.LXW_FALSE,
        .output_buffer = null,
        .output_buffer_size = null,
    };

    // Create a new workbook with options
    const workbook = xlsxwriter.workbook_new_opt("zig-constant_memory.xlsx", &options);
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    var row: u32 = 0;
    while (row < row_max) : (row += 1) {
        var col: u16 = 0;
        while (col < col_max) : (col += 1) {
            _ = xlsxwriter.worksheet_write_number(worksheet, row, col, 123.45, null);
        }
    }

    _ = xlsxwriter.workbook_close(workbook);
}
