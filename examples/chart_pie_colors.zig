//
// An example of creating an Excel pie chart with user defined colors using
// the libxlsxwriter library.
//
// In general formatting is applied to an entire series in a chart. However,
// it is occasionally required to format individual points in a series. In
// particular this is required for Pie/Doughnut charts where each segment is
// represented by a point.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_pie_colors.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
