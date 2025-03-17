//
// Example of writing dates and times in Excel using different date formats.
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
        .month = 1,
        .day = 23,
        .hour = 12,
        .min = 30,
        .sec = 5.123,
    };
    var row: u32 = 0;
    const col: u16 = 0;

    // Examples date and time formats. In the output file compare how changing
    // the format strings changes the appearance of the date.
    const date_formats = [_][]const u8{
        "dd/mm/yy",
        "mm/dd/yy",
        "dd m yy",
        "d mm yy",
        "d mmm yy",
        "d mmmm yy",
        "d mmmm yyy",
        "d mmmm yyyy",
        "dd/mm/yy hh:mm",
        "dd/mm/yy hh:mm:ss",
        "dd/mm/yy hh:mm:ss.000",
        "hh:mm",
        "hh:mm:ss",
        "hh:mm:ss.000",
    };

    const workbook = xlsxwriter.workbook_new("zig-dates_and_times04.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a bold format.
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Write the column headers.
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col, "Formatted date", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 1, "Format", bold);

    // Widen the first column to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 1, 20, null);

    // Write the same date and time using each of the above formats.
    for (date_formats) |date_format| {
        row += 1;

        // Create a format for the date or time.
        const format = xlsxwriter.workbook_add_format(workbook);
        _ = xlsxwriter.format_set_num_format(format, date_format.ptr);
        _ = xlsxwriter.format_set_align(format, xlsxwriter.LXW_ALIGN_LEFT);

        // Write the datetime with each format.
        _ = xlsxwriter.worksheet_write_datetime(worksheet, row, col, &datetime, format);

        // Also write the format string for comparison.
        _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 1, date_format.ptr, null);
    }

    _ = xlsxwriter.workbook_close(workbook);
}
