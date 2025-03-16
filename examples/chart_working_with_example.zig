//
// An example of a simple Excel chart using the libxlsxwriter library. This
// example is used in the "Working with Charts" section of the docs.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_working_with_example.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Write some data for the chart
    _ = xlsxwriter.worksheet_write_number(worksheet, 0, 0, 10, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 40, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 0, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 0, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 0, 10, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 5, 0, 50, null);

    // Create a chart object
    const chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_LINE);

    // Configure the chart
    const series = xlsxwriter.chart_add_series(chart, null, "Sheet1!$A$1:$A$6");
    _ = series; // Used in other examples

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 0, 2, chart);

    _ = xlsxwriter.workbook_close(workbook);
}
