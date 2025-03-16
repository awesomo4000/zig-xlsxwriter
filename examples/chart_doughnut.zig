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

// Write some data to the worksheet
fn write_worksheet_data(worksheet: *xlsxwriter.lxw_worksheet, bold: *xlsxwriter.lxw_format) void {
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Category", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "Glazed", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "Chocolate", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 0, "Cream", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Values", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 1, 35, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 1, 15, null);
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_doughnut.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);
    var chart: *xlsxwriter.lxw_chart = undefined;
    var series: *xlsxwriter.lxw_chart_series = undefined;

    // Add a bold format to use to highlight the header cells
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Write some data for the chart
    write_worksheet_data(worksheet, bold);

    // Chart 1: Create a simple doughnut chart
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_DOUGHNUT);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "Doughnut sales data");

    // Add a chart title
    _ = xlsxwriter.chart_title_set_name(chart, "Popular Doughnut Types");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart, 10);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 3, chart);

    // Chart 2: Create a doughnut chart with user defined segment colors
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_DOUGHNUT);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "Doughnut sales data");

    // Add a chart title
    _ = xlsxwriter.chart_title_set_name(chart, "Doughnut Chart with user defined colors");

    // Add fills for use in the chart
    var fill1 = xlsxwriter.lxw_chart_fill{ .color = 0xFA58D0 };
    var fill2 = xlsxwriter.lxw_chart_fill{ .color = 0x61210B };
    var fill3 = xlsxwriter.lxw_chart_fill{ .color = 0xF5F6CE };

    // Add some points with the above fills
    var point1 = xlsxwriter.lxw_chart_point{ .fill = &fill1 };
    var point2 = xlsxwriter.lxw_chart_point{ .fill = &fill2 };
    var point3 = xlsxwriter.lxw_chart_point{ .fill = &fill3 };

    // Create an array of the point objects
    const points_array = [_]*xlsxwriter.lxw_chart_point{
        &point1,
        &point2,
        &point3,
    };

    // Create a null-terminated array of pointers as required by the C API
    var points_ptrs: [4][*c]xlsxwriter.lxw_chart_point = undefined;
    for (points_array, 0..) |point, i| {
        points_ptrs[i] = point;
    }
    points_ptrs[3] = null;

    // Add/override the points/segments of the chart
    _ = xlsxwriter.chart_series_set_points(series, @ptrCast(&points_ptrs));

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 17, 3, chart);

    // Chart 3: Create a Doughnut chart with rotation of the segments
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_DOUGHNUT);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "Doughnut sales data");

    // Add a chart title
    _ = xlsxwriter.chart_title_set_name(chart, "Doughnut Chart with segment rotation");

    // Change the angle/rotation of the first segment
    _ = xlsxwriter.chart_set_rotation(chart, 90);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 33, 3, chart);

    // Chart 4: Create a Doughnut chart with user defined hole size and other options
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_DOUGHNUT);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "Doughnut sales data");

    // Add a chart title
    _ = xlsxwriter.chart_title_set_name(chart, "Doughnut Chart with options applied.");

    // Add/override the points/segments defined in Chart 2
    _ = xlsxwriter.chart_series_set_points(series, @ptrCast(&points_ptrs));

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart, 26);

    // Change the angle/rotation of the first segment
    _ = xlsxwriter.chart_set_rotation(chart, 28);

    // Change the hole size
    _ = xlsxwriter.chart_set_hole_size(chart, 33);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 49, 3, chart);

    _ = xlsxwriter.workbook_close(workbook);
}
