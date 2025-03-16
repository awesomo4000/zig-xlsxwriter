//
// Example of writing dates and times in Excel using an lxw_datetime struct
// and date formatting.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // A datetime to display.
    var datetime = xlsxwriter.lxw_datetime{
        .year = 2013,
        .month = 2,
        .day = 28,
        .hour = 12,
        .min = 0,
        .sec = 0.0,
    };

    const workbook = xlsxwriter.workbook_new("zig-dates_and_times02.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a format with date formatting.
    const format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(format, "mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write the datetime without formatting.
    _ = xlsxwriter.worksheet_write_datetime(worksheet, 0, 0, &datetime, null); // 41333.5

    // Write the datetime with formatting.
    _ = xlsxwriter.worksheet_write_datetime(worksheet, 1, 0, &datetime, format); // Feb 28 2013 12:00 PM

    _ = xlsxwriter.workbook_close(workbook);
}
