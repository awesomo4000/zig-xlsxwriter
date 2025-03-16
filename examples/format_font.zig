// Example of writing some data with font formatting to a simple Excel
// file using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    // Create a new workbook.
    const workbook = xlsxwriter.workbook_new("zig-format_font.xlsx");

    // Add a worksheet.
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Widen the first column to make the text clearer.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Add some formats.
    const format1 = xlsxwriter.workbook_add_format(workbook);
    const format2 = xlsxwriter.workbook_add_format(workbook);
    const format3 = xlsxwriter.workbook_add_format(workbook);

    // Set the bold property for format 1.
    _ = xlsxwriter.format_set_bold(format1);

    // Set the italic property for format 2.
    _ = xlsxwriter.format_set_italic(format2);

    // Set the bold and italic properties for format 3.
    _ = xlsxwriter.format_set_bold(format3);
    _ = xlsxwriter.format_set_italic(format3);

    // Write some formatted strings.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "This is bold", format1);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "This is italic", format2);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "Bold and italic", format3);

    // Close the workbook, save the file and free any memory.
    _ = xlsxwriter.workbook_close(workbook);
}
