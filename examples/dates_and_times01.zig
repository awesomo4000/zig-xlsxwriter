//
// Example of writing a dates and time in Excel using a number with date
// formatting. This demonstrates that dates and times in Excel are just
// formatted real numbers.
//
// An easier approach using a lxw_datetime struct is shown in example
// dates_and_times02.c.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // A number to display as a date.
    const number = 41333.5;

    const workbook = xlsxwriter.workbook_new("zig-dates_and_times01.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a format with date formatting.
    const format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(format, "mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write the number without formatting.
    _ = xlsxwriter.worksheet_write_number(worksheet, 0, 0, number, null); // 41333.5

    // Write the number with formatting. Note: the worksheet_write_datetime()
    // or worksheet_write_unixtime() functions are preferable for writing
    // dates and times. This is for demonstration purposes only.
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, number, format); // Feb 28 2013 12:00 PM

    const result = xlsxwriter.workbook_close(workbook);
    if (result != xlsxwriter.LXW_NO_ERROR) {
        std.debug.print("Error closing workbook: {d}\n", .{result});
        return error.WorkbookCloseFailed;
    }
}
