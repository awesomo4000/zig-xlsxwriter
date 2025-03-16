//
// Example of how use libxlsxwriter to generate Excel outlines and grouping.
//
// Excel allows you to group rows or columns so that they can be hidden or
// displayed with a single mouse click. This feature is referred to as
// outlines.
//
// Outlines can reduce complex data down to a few salient sub-totals or
// summaries.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-outline.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
