//
// An example of creating Excel scatter charts using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Write some data to the worksheet.
fn write_worksheet_data(worksheet: *xlsxwriter.lxw_worksheet, bold: *xlsxwriter.lxw_format) void {
    const data = [_][3]f64{
        // Three columns of data
        [_]f64{ 2, 10, 30 },
        [_]f64{ 3, 40, 60 },
        [_]f64{ 4, 50, 70 },
        [_]f64{ 5, 20, 50 },
        [_]f64{ 6, 10, 40 },
        [_]f64{ 7, 50, 30 },
    };

    // Write the column headers
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Number", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Batch 1", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "Batch 2", bold);

    // Write the example data
    for (data, 0..) |row, i| {
        for (row, 0..) |value, j| {
            _ = xlsxwriter.worksheet_write_number(worksheet, @intCast(i + 1), @intCast(j), value, null);
        }
    }
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_scatter.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Write some data for the chart
    write_worksheet_data(worksheet, bold);

    // Chart 1: Create a scatter chart
    const chart1 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_SCATTER);

    // Add the first series to the chart
    var series = xlsxwriter.chart_add_series(chart1, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart
    series = xlsxwriter.chart_add_series(chart1, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels
    _ = xlsxwriter.chart_title_set_name(chart1, "Results of sample analysis");
    _ = xlsxwriter.chart_axis_set_name(chart1.*.x_axis, "Test number");
    _ = xlsxwriter.chart_axis_set_name(chart1.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart1, 11);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 4, chart1);

    // Chart 2: Create a scatter chart with straight lines and markers
    const chart2 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart2, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart
    series = xlsxwriter.chart_add_series(chart2, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels
    _ = xlsxwriter.chart_title_set_name(chart2, "Results of sample analysis");
    _ = xlsxwriter.chart_axis_set_name(chart2.*.x_axis, "Test number");
    _ = xlsxwriter.chart_axis_set_name(chart2.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart2, 12);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 17, 4, chart2);

    // Chart 3: Create a scatter chart with straight lines
    const chart3 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_SCATTER_STRAIGHT);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart3, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart
    series = xlsxwriter.chart_add_series(chart3, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels
    _ = xlsxwriter.chart_title_set_name(chart3, "Results of sample analysis");
    _ = xlsxwriter.chart_axis_set_name(chart3.*.x_axis, "Test number");
    _ = xlsxwriter.chart_axis_set_name(chart3.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart3, 13);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 33, 4, chart3);

    // Chart 4: Create a scatter chart with smooth lines and markers
    const chart4 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart4, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart
    series = xlsxwriter.chart_add_series(chart4, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels
    _ = xlsxwriter.chart_title_set_name(chart4, "Results of sample analysis");
    _ = xlsxwriter.chart_axis_set_name(chart4.*.x_axis, "Test number");
    _ = xlsxwriter.chart_axis_set_name(chart4.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart4, 14);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 49, 4, chart4);

    // Chart 5: Create a scatter chart with smooth lines
    const chart5 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_SCATTER_SMOOTH);

    // Add the first series to the chart
    series = xlsxwriter.chart_add_series(chart5, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart
    series = xlsxwriter.chart_add_series(chart5, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Set the name for the series instead of the default "Series 2"
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels
    _ = xlsxwriter.chart_title_set_name(chart5, "Results of sample analysis");
    _ = xlsxwriter.chart_axis_set_name(chart5.*.x_axis, "Test number");
    _ = xlsxwriter.chart_axis_set_name(chart5.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style
    _ = xlsxwriter.chart_set_style(chart5, 15);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 65, 4, chart5);

    _ = xlsxwriter.workbook_close(workbook);
}
