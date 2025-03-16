//
// Example of how to hide a worksheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-hide_sheet.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Hide Sheet2. It won't be visible until it is unhidden in Excel.
    _ = xlsxwriter.worksheet_hide(worksheet2);

    _ = xlsxwriter.worksheet_write_string(worksheet1, 0, 0, "Sheet2 is hidden", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 0, "Now it's my turn to find you!", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 0, "Sheet2 is hidden", null);

    // Make the first column wider to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet1, 0, 0, 30, null);
    _ = xlsxwriter.worksheet_set_column(worksheet2, 0, 0, 30, null);
    _ = xlsxwriter.worksheet_set_column(worksheet3, 0, 0, 30, null);

    _ = xlsxwriter.workbook_close(workbook);
}
