// Example of writing some data to a simple Excel file using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-hello.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Hello", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 123, null);

    _ = xlsxwriter.workbook_close(workbook);
}
