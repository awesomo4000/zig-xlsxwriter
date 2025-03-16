//
// An example of merging cells containing a rich string using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-merge_rich_string.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
