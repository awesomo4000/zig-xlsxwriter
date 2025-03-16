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

    // Write data for the chart
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Pass", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "Fail", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 0, 1, 90, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 10, null);

    // Create a pie chart
    const chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_PIE);

    // Add the data series
    const series = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$1:$A$2",
        "=Sheet1!$B$1:$B$2",
    );

    // Create fills for chart segments
    var red_fill = xlsxwriter.lxw_chart_fill{
        .color = xlsxwriter.LXW_COLOR_RED,
    };
    var green_fill = xlsxwriter.lxw_chart_fill{
        .color = xlsxwriter.LXW_COLOR_GREEN,
    };

    // Create points with fills
    var red_point = xlsxwriter.lxw_chart_point{
        .fill = &red_fill,
    };
    var green_point = xlsxwriter.lxw_chart_point{
        .fill = &green_fill,
    };

    // Create array of points (null terminated)
    var points = [_]?*xlsxwriter.lxw_chart_point{
        &green_point,
        &red_point,
        null,
    };

    // Set the points on the series
    _ = xlsxwriter.chart_series_set_points(series, &points);

    // Insert chart into worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 3, chart);

    _ = xlsxwriter.workbook_close(workbook);
}
