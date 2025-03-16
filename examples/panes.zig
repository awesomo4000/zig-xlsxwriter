const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    var row: u32 = 0;
    var col: u16 = 0;

    // Create a new workbook and add some worksheets.
    const workbook = xlsxwriter.workbook_new("zig-panes.xlsx");

    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, "Panes 1");
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, "Panes 2");
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, "Panes 3");
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, "Panes 4");

    // Set up some formatting and text to highlight the panes.
    const header = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_align(header, xlsxwriter.LXW_ALIGN_CENTER);
    _ = xlsxwriter.format_set_align(header, xlsxwriter.LXW_ALIGN_VERTICAL_CENTER);
    _ = xlsxwriter.format_set_fg_color(header, 0xD7E4BC);
    _ = xlsxwriter.format_set_bold(header);
    _ = xlsxwriter.format_set_border(header, xlsxwriter.LXW_BORDER_THIN);

    const center = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_align(center, xlsxwriter.LXW_ALIGN_CENTER);

    //
    // Example 1. Freeze pane on the top row.
    //
    _ = xlsxwriter.worksheet_freeze_panes(worksheet1, 1, 0);

    // Some sheet formatting.
    _ = xlsxwriter.worksheet_set_column(worksheet1, 0, 8, 16, null);
    _ = xlsxwriter.worksheet_set_row(worksheet1, 0, 20, null);
    _ = xlsxwriter.worksheet_set_selection(worksheet1, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 0;
    while (col < 9) : (col += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet1, 0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 100) : (row += 1) {
        col = 0;
        while (col < 9) : (col += 1) {
            _ = xlsxwriter.worksheet_write_number(worksheet1, row, col, @floatFromInt(row + 1), center);
        }
    }

    //
    // Example 2. Freeze pane on the left column.
    //
    _ = xlsxwriter.worksheet_freeze_panes(worksheet2, 0, 1);

    // Some sheet formatting.
    _ = xlsxwriter.worksheet_set_column(worksheet2, 0, 0, 16, null);
    _ = xlsxwriter.worksheet_set_selection(worksheet2, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    row = 0;
    while (row < 50) : (row += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet2, row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = xlsxwriter.worksheet_write_number(worksheet2, row, col, @floatFromInt(col), center);
        }
    }

    //
    // Example 3. Freeze pane on the top row and left column.
    //
    _ = xlsxwriter.worksheet_freeze_panes(worksheet3, 1, 1);

    // Some sheet formatting.
    _ = xlsxwriter.worksheet_set_column(worksheet3, 0, 25, 16, null);
    _ = xlsxwriter.worksheet_set_row(worksheet3, 0, 20, null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 0, "", header);
    _ = xlsxwriter.worksheet_set_selection(worksheet3, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet3, 0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet3, row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = xlsxwriter.worksheet_write_number(worksheet3, row, col, @floatFromInt(col), center);
        }
    }

    //
    // Example 4. Split pane on the top row and left column.
    //
    // The divisions must be specified in terms of row and column dimensions.
    // The default row height is 15 and the default column width is 8.43
    //
    _ = xlsxwriter.worksheet_split_panes(worksheet4, 15, 8.43);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet4, 0, col, "Scroll", center);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        _ = xlsxwriter.worksheet_write_string(worksheet4, row, 0, "Scroll", center);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = xlsxwriter.worksheet_write_number(worksheet4, row, col, @floatFromInt(col), center);
        }
    }

    _ = xlsxwriter.workbook_close(workbook);
}
