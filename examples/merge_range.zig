const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-merge_range.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);
    const merge_format = xlsxwriter.workbook_add_format(workbook);

    // Configure a format for the merged range.
    _ = xlsxwriter.format_set_align(merge_format, xlsxwriter.LXW_ALIGN_CENTER);
    _ = xlsxwriter.format_set_align(merge_format, xlsxwriter.LXW_ALIGN_VERTICAL_CENTER);
    _ = xlsxwriter.format_set_bold(merge_format);
    _ = xlsxwriter.format_set_bg_color(merge_format, xlsxwriter.LXW_COLOR_YELLOW);
    _ = xlsxwriter.format_set_border(merge_format, xlsxwriter.LXW_BORDER_THIN);

    // Increase the cell size of the merged cells to highlight the formatting.
    _ = xlsxwriter.worksheet_set_column(worksheet, 1, 3, 12, null);
    _ = xlsxwriter.worksheet_set_row(worksheet, 3, 30, null);
    _ = xlsxwriter.worksheet_set_row(worksheet, 6, 30, null);
    _ = xlsxwriter.worksheet_set_row(worksheet, 7, 30, null);

    // Merge 3 cells.
    _ = xlsxwriter.worksheet_merge_range(worksheet, 3, 1, 3, 3, "Merged Range", merge_format);

    // Merge 3 cells over two rows.
    _ = xlsxwriter.worksheet_merge_range(worksheet, 6, 1, 7, 3, "Merged Range", merge_format);

    _ = xlsxwriter.workbook_close(workbook);
}
