//
// An example of how to hide rows and columns using the libxlsxwriter
// library.
//
// In order to hide rows without setting each one, (of approximately 1 million
// rows), Excel uses an optimization to hide all rows that don't have data. In
// Libxlsxwriter we replicate that using the worksheet_set_default_row()
// function.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-hide_row_col.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
