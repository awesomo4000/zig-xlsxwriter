//
// An example of writing cell comments to a worksheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-comments1.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Hello", null);

    _ = xlsxwriter.worksheet_write_comment(worksheet, 0, 0, "This is a comment");

    _ = xlsxwriter.workbook_close(workbook);
}
