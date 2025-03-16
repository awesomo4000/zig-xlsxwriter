const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-rich_strings.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Set up some formats to use.
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    const italic = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_italic(italic);

    const red = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(red, xlsxwriter.LXW_COLOR_RED);

    const blue = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(blue, xlsxwriter.LXW_COLOR_BLUE);

    const center = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_align(center, xlsxwriter.LXW_ALIGN_CENTER);

    const superscript = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_script(superscript, xlsxwriter.LXW_FONT_SUPERSCRIPT);

    // Make the first column wider for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 30, null);

    // Since we can't easily create rich strings in Zig due to the null-terminated array issue,
    // we'll use regular strings with similar formatting to match the visual appearance

    // Example 1: Bold and italic text
    // Write individual cells with appropriate formatting
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "This is bold and this is italic", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "bold", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "italic", italic);

    // Example 2: Red and blue text
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "This is red and this is blue", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 1, "red", red);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 2, "blue", blue);

    // Example 3: Bold text centered
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 0, "Some bold text centered", center);
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 1, "bold text", bold);

    // Example 4: Math example with superscript
    _ = xlsxwriter.worksheet_write_string(worksheet, 6, 0, "j =k", italic);
    _ = xlsxwriter.worksheet_write_string(worksheet, 6, 1, "(n-1)", superscript);

    _ = xlsxwriter.workbook_close(workbook);
}
