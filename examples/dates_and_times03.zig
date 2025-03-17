//
// Example of writing dates and times in Excel using a Unix datetime and date
// formatting.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-dates_and_times03.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a format with date formatting.
    const format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(format, "mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write some Unix datetimes with formatting.

    // 1970-01-01. The Unix epoch.
    _ = xlsxwriter.worksheet_write_unixtime(worksheet, 0, 0, 0, format);

    // 2000-01-01.
    _ = xlsxwriter.worksheet_write_unixtime(worksheet, 1, 0, 1577836800, format);

    // 1900-01-01.
    _ = xlsxwriter.worksheet_write_unixtime(worksheet, 2, 0, -2208988800, format);

    _ = xlsxwriter.workbook_close(workbook);
}
