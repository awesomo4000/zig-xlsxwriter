//
// A simple formatting example that demonstrates how to add diagonal
// cell borders using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-diagonal_border.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add some diagonal border formats.
    const format1 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_diag_type(format1, xlsxwriter.LXW_DIAGONAL_BORDER_UP);

    const format2 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_diag_type(format2, xlsxwriter.LXW_DIAGONAL_BORDER_DOWN);

    const format3 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_diag_type(format3, xlsxwriter.LXW_DIAGONAL_BORDER_UP_DOWN);

    const format4 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_diag_type(format4, xlsxwriter.LXW_DIAGONAL_BORDER_UP_DOWN);
    _ = xlsxwriter.format_set_diag_border(format4, xlsxwriter.LXW_BORDER_HAIR);
    _ = xlsxwriter.format_set_diag_color(format4, xlsxwriter.LXW_COLOR_RED);

    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 1, "Text", format1);
    _ = xlsxwriter.worksheet_write_string(worksheet, 5, 1, "Text", format2);
    _ = xlsxwriter.worksheet_write_string(worksheet, 8, 1, "Text", format3);
    _ = xlsxwriter.worksheet_write_string(worksheet, 11, 1, "Text", format4);

    _ = xlsxwriter.workbook_close(workbook);
}
