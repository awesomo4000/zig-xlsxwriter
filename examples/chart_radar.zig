//
// An example of creating Excel radar charts using the libxlsxwriter library.
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
        [_]f64{ 2, 30, 25 },
        [_]f64{ 3, 60, 40 },
        [_]f64{ 4, 70, 50 },
        [_]f64{ 5, 50, 30 },
        [_]f64{ 6, 40, 50 },
        [_]f64{ 7, 30, 40 },
    };

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Number", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Batch 1", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "Batch 2", bold);

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            _ = xlsxwriter.worksheet_write_number(
                worksheet,
                @intCast(row_idx + 1),
                @intCast(col_idx),
                value,
                null,
            );
        }
    }
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_radar.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells.
    const bold = xlsxwriter.workbook_add_format(workbook);
    xlsxwriter.format_set_bold(bold);

    // Write some data for the chart.
    write_worksheet_data(worksheet, bold);

    // Chart 1: Create a radar chart.
    const chart1 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_RADAR);

    // Add the first series to the chart.
    var series = xlsxwriter.chart_add_series(
        chart1,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the series instead of the default "Series 1".
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined.
    series = xlsxwriter.chart_add_series(chart1, null, null);

    // Configure the series using a syntax that is easier to define programmatically.
    _ = xlsxwriter.chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
    _ = xlsxwriter.chart_series_set_values(series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
    _ = xlsxwriter.chart_series_set_name_range(series, "Sheet1", 0, 2); // "=Sheet1!$C$1"

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart1, "Results of sample analysis");

    // Set an Excel chart style.
    _ = xlsxwriter.chart_set_style(chart1, 11);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 1, 4, chart1);

    // Chart 2: Create a radar chart with markers.
    const chart2 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_RADAR_WITH_MARKERS);

    // Add the first series to the chart.
    series = xlsxwriter.chart_add_series(
        chart2,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the series instead of the default "Series 1".
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart.
    series = xlsxwriter.chart_add_series(
        chart2,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );

    // Set the name for the series instead of the default "Series 2".
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart2, "Results of sample analysis");

    // Set an Excel chart style.
    _ = xlsxwriter.chart_set_style(chart2, 12);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 17, 4, chart2);

    // Chart 3: Create a filled radar chart.
    const chart3 = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_RADAR_FILLED);

    // Add the first series to the chart.
    series = xlsxwriter.chart_add_series(
        chart3,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the series instead of the default "Series 1".
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart.
    series = xlsxwriter.chart_add_series(
        chart3,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );

    // Set the name for the series instead of the default "Series 2".
    _ = xlsxwriter.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart3, "Results of sample analysis");

    // Set an Excel chart style.
    _ = xlsxwriter.chart_set_style(chart3, 13);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 33, 4, chart3);

    _ = xlsxwriter.workbook_close(workbook);
}
