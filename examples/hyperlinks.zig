//
// Example of writing urls/hyperlinks with the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // Create a new workbook
    const workbook = xlsxwriter.workbook_new("zig-hyperlinks.xlsx");

    // Add a worksheet
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Get the default url format (used in the overwriting examples below)
    const url_format = xlsxwriter.workbook_get_default_url_format(workbook);

    // Create a user defined link format
    const red_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_underline(red_format, xlsxwriter.LXW_UNDERLINE_SINGLE);
    _ = xlsxwriter.format_set_font_color(red_format, xlsxwriter.LXW_COLOR_RED);

    // Widen the first column to make the text clearer
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 30, null);

    // Write a hyperlink. A default blue underline will be used if the format is NULL
    _ = xlsxwriter.worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", null);

    // Write a hyperlink but overwrite the displayed string. Note, we need to
    // specify the format for the string to match the default hyperlink
    _ = xlsxwriter.worksheet_write_url(worksheet, 2, 0, "http://libxlsxwriter.github.io", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "Read the documentation.", url_format);

    // Write a hyperlink with a different format
    _ = xlsxwriter.worksheet_write_url(worksheet, 4, 0, "http://libxlsxwriter.github.io", red_format);

    // Write a mail hyperlink
    _ = xlsxwriter.worksheet_write_url(worksheet, 6, 0, "mailto:jmcnamara@cpan.org", null);

    // Write a mail hyperlink and overwrite the displayed string. We again
    // specify the format for the string to match the default hyperlink
    _ = xlsxwriter.worksheet_write_url(worksheet, 8, 0, "mailto:jmcnamara@cpan.org", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 8, 0, "Drop me a line.", url_format);

    // Close the workbook, save the file and free any memory
    _ = xlsxwriter.workbook_close(workbook);
}
