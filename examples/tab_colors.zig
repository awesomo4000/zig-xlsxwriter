//
// Example of how to set Excel worksheet tab colors using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-tab_colors.xlsx");

    // Set up some worksheets.
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Set the tab colors.
    _ = xlsxwriter.worksheet_set_tab_color(worksheet1, xlsxwriter.LXW_COLOR_RED);
    _ = xlsxwriter.worksheet_set_tab_color(worksheet2, xlsxwriter.LXW_COLOR_GREEN);
    _ = xlsxwriter.worksheet_set_tab_color(worksheet3, 0xFF9900); // Orange.

    // worksheet4 will have the default color.
    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 0, "Hello", null);

    _ = xlsxwriter.workbook_close(workbook);
}
