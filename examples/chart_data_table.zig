//
// An example of creating Excel column charts with data tables using the
// libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_data_table.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
