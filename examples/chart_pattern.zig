//
// An example of a simple Excel chart with patterns using the libxlsxwriter
// library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_pattern.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells.
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Write some data for the chart.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Shingle", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 105, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 0, 150, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 0, 130, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 0, 90, null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Brick", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 1, 120, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 1, 100, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 1, 110, null);

    // Create a chart object.
    const chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Configure the chart.
    const series1 = xlsxwriter.chart_add_series(chart, null, "Sheet1!$A$2:$A$5");
    const series2 = xlsxwriter.chart_add_series(chart, null, "Sheet1!$B$2:$B$5");

    _ = xlsxwriter.chart_series_set_name(series1, "=Sheet1!$A$1");
    _ = xlsxwriter.chart_series_set_name(series2, "=Sheet1!$B$1");

    _ = xlsxwriter.chart_title_set_name(chart, "Cladding types");
    _ = xlsxwriter.chart_axis_set_name(chart.*.x_axis, "Region");
    _ = xlsxwriter.chart_axis_set_name(chart.*.y_axis, "Number of houses");

    // Configure and add the chart series patterns.
    var pattern1 = xlsxwriter.lxw_chart_pattern{
        .type = xlsxwriter.LXW_CHART_PATTERN_SHINGLE,
        .fg_color = 0x804000,
        .bg_color = 0xC68C53,
    };

    var pattern2 = xlsxwriter.lxw_chart_pattern{
        .type = xlsxwriter.LXW_CHART_PATTERN_HORIZONTAL_BRICK,
        .fg_color = 0xB30000,
        .bg_color = 0xFF6666,
    };

    _ = xlsxwriter.chart_series_set_pattern(series1, &pattern1);
    _ = xlsxwriter.chart_series_set_pattern(series2, &pattern2);

    // Configure and set the chart series borders.
    var line1 = xlsxwriter.lxw_chart_line{
        .color = 0x804000,
        .none = 0,
        .width = 0,
        .dash_type = 0,
        .transparency = 0,
    };

    var line2 = xlsxwriter.lxw_chart_line{
        .color = 0xb30000,
        .none = 0,
        .width = 0,
        .dash_type = 0,
        .transparency = 0,
    };

    _ = xlsxwriter.chart_series_set_line(series1, &line1);
    _ = xlsxwriter.chart_series_set_line(series2, &line2);

    // Widen the gap between the series/categories.
    _ = xlsxwriter.chart_set_series_gap(chart, 70);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 3, chart);

    _ = xlsxwriter.workbook_close(workbook);
}
