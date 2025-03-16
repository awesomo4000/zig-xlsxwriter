//
// A simple example of some of the features of the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // Create a new workbook and add a worksheet.
    const workbook = xlsxwriter.workbook_new("zig-demo.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a format.
    const format = xlsxwriter.workbook_add_format(workbook);

    // Set the bold property for the format
    _ = xlsxwriter.format_set_bold(format);

    // Change the column width for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write some simple text.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Hello", null);

    // Text with formatting.
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "World", format);

    // Write some numbers.
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 0, 123, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 0, 123.456, null);

    // Insert an image.
    _ = xlsxwriter.worksheet_insert_image(worksheet, 1, 2, "logo.png");

    _ = xlsxwriter.workbook_close(workbook);
}
