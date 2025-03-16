//
// An example of creating an Excel doughnut chart using the libxlsxwriter library.
//
// The demo also shows how to set segment colors. It is possible to define
// chart colors for most types of libxlsxwriter charts via the series
// formatting functions. However, Pie/Doughnut charts are a special case since
// each segment is represented as a point so it is necessary to assign
// formatting to each point in the series.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_doughnut.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
