//
// A demo of an various Excel chart data label features that are available via
// a libxlsxwriter chart.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_data_labels.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells.
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Some chart positioning options.
    var options = xlsxwriter.lxw_chart_options{
        .x_offset = 25,
        .y_offset = 10,
    };

    // Write some data for the chart.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Number", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 2, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 0, 3, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 0, 4, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 0, 5, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 5, 0, 6, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 6, 0, 7, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Data", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 1, 10, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 1, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 1, 30, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 5, 1, 40, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 6, 1, 30, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "Text", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 2, "Jan", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 2, "Feb", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 2, "Mar", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 2, "Apr", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 5, 2, "May", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 6, 2, "Jun", null);

    // Chart 1. Example with standard data labels.
    var chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Chart with standard data labels");

    // Add a data series to the chart.
    const series = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 1, 3, chart, &options);

    // Chart 2. Example with value and category data labels.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Category and Value data labels");

    // Add a data series to the chart.
    const series2 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series2);

    // Turn on Value and Category labels.
    _ = xlsxwriter.chart_series_set_labels_options(series2, 0, 1, 1);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 17, 3, chart, &options);

    // Chart 3. Example with standard data labels with different font.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Data labels with user defined font");

    // Add a data series to the chart.
    const series3 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series3);

    var font1 = xlsxwriter.lxw_chart_font{
        .bold = 1,
        .color = xlsxwriter.LXW_COLOR_RED,
        .rotation = -30,
    };
    _ = xlsxwriter.chart_series_set_labels_font(series3, &font1);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 33, 3, chart, &options);

    // Chart 4. Example with standard data labels and formatting.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Data labels with formatting");

    // Add a data series to the chart.
    const series4 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series4);

    // Set the border/line and fill for the data labels.
    var line1 = xlsxwriter.lxw_chart_line{ .color = xlsxwriter.LXW_COLOR_RED };
    var fill1 = xlsxwriter.lxw_chart_fill{ .color = xlsxwriter.LXW_COLOR_YELLOW };

    _ = xlsxwriter.chart_series_set_labels_line(series4, &line1);
    _ = xlsxwriter.chart_series_set_labels_fill(series4, &fill1);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 49, 3, chart, &options);

    // Chart 5. Example with custom string data labels.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Chart with custom string data labels");

    // Add a data series to the chart.
    const series5 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series5);

    // Create some custom labels.
    var data_label5_1 = xlsxwriter.lxw_chart_data_label{ .value = "Amy" };
    var data_label5_2 = xlsxwriter.lxw_chart_data_label{ .value = "Bea" };
    var data_label5_3 = xlsxwriter.lxw_chart_data_label{ .value = "Eva" };
    var data_label5_4 = xlsxwriter.lxw_chart_data_label{ .value = "Fay" };
    var data_label5_5 = xlsxwriter.lxw_chart_data_label{ .value = "Liv" };
    var data_label5_6 = xlsxwriter.lxw_chart_data_label{ .value = "Una" };

    // Create an array of label pointers.
    var data_labels5: [7]?*xlsxwriter.lxw_chart_data_label = undefined;
    data_labels5[0] = &data_label5_1;
    data_labels5[1] = &data_label5_2;
    data_labels5[2] = &data_label5_3;
    data_labels5[3] = &data_label5_4;
    data_labels5[4] = &data_label5_5;
    data_labels5[5] = &data_label5_6;
    data_labels5[6] = null;

    // Set the custom labels.
    _ = xlsxwriter.chart_series_set_labels_custom(series5, &data_labels5);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 65, 3, chart, &options);

    // Chart 6. Example with custom data labels from cells.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Chart with custom data labels from cells");

    // Add a data series to the chart.
    const series6 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series6);

    // Create some custom labels.
    var data_label6_1 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$2" };
    var data_label6_2 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$3" };
    var data_label6_3 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$4" };
    var data_label6_4 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$5" };
    var data_label6_5 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$6" };
    var data_label6_6 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$7" };

    // Create an array of label pointers.
    var data_labels6: [7]?*xlsxwriter.lxw_chart_data_label = undefined;
    data_labels6[0] = &data_label6_1;
    data_labels6[1] = &data_label6_2;
    data_labels6[2] = &data_label6_3;
    data_labels6[3] = &data_label6_4;
    data_labels6[4] = &data_label6_5;
    data_labels6[5] = &data_label6_6;
    data_labels6[6] = null;

    // Set the custom labels.
    _ = xlsxwriter.chart_series_set_labels_custom(series6, &data_labels6);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 81, 3, chart, &options);

    // Chart 7. Example with custom and default data labels.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Mixed custom and default data labels");

    // Add a data series to the chart.
    const series7 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    var font2 = xlsxwriter.lxw_chart_font{ .color = xlsxwriter.LXW_COLOR_RED };

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series7);

    // Create some custom labels.
    var data_label7_1 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$2", .font = &font2 };
    var data_label7_2 = xlsxwriter.lxw_chart_data_label{};
    var data_label7_3 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$4", .font = &font2 };
    var data_label7_4 = xlsxwriter.lxw_chart_data_label{ .value = "=Sheet1!$C$5", .font = &font2 };

    // Create an array of label pointers.
    var data_labels7: [5]?*xlsxwriter.lxw_chart_data_label = undefined;
    data_labels7[0] = &data_label7_1;
    data_labels7[1] = &data_label7_2;
    data_labels7[2] = &data_label7_3;
    data_labels7[3] = &data_label7_4;
    data_labels7[4] = null;

    // Set the custom labels.
    _ = xlsxwriter.chart_series_set_labels_custom(series7, &data_labels7);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 97, 3, chart, &options);

    // Chart 8. Example with deleted/hidden data labels.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Chart with deleted data labels");

    // Add a data series to the chart.
    const series8 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series8);

    // Create some custom labels.
    var hide8 = xlsxwriter.lxw_chart_data_label{ .hide = 1 };
    var keep8 = xlsxwriter.lxw_chart_data_label{ .hide = 0 };

    // Create an array of label pointers.
    var data_labels8: [7]?*xlsxwriter.lxw_chart_data_label = undefined;
    data_labels8[0] = &hide8;
    data_labels8[1] = &keep8;
    data_labels8[2] = &hide8;
    data_labels8[3] = &hide8;
    data_labels8[4] = &keep8;
    data_labels8[5] = &hide8;
    data_labels8[6] = null;

    // Set the custom labels.
    _ = xlsxwriter.chart_series_set_labels_custom(series8, &data_labels8);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 113, 3, chart, &options);

    // Chart 9. Example with custom string data labels and formatting.
    chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_COLUMN);

    // Add a chart title.
    _ = xlsxwriter.chart_title_set_name(chart, "Chart with custom labels and formatting");

    // Add a data series to the chart.
    const series9 = xlsxwriter.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Add the series data labels.
    _ = xlsxwriter.chart_series_set_labels(series9);

    // Set the border/line and fill for the data labels.
    var line2 = xlsxwriter.lxw_chart_line{ .color = xlsxwriter.LXW_COLOR_RED };
    var fill2 = xlsxwriter.lxw_chart_fill{ .color = xlsxwriter.LXW_COLOR_YELLOW };
    var line3 = xlsxwriter.lxw_chart_line{ .color = xlsxwriter.LXW_COLOR_BLUE };
    var fill3 = xlsxwriter.lxw_chart_fill{ .color = xlsxwriter.LXW_COLOR_GREEN };

    // Create some custom labels.
    var data_label9_1 = xlsxwriter.lxw_chart_data_label{ .value = "Amy", .line = &line3 };
    var data_label9_2 = xlsxwriter.lxw_chart_data_label{ .value = "Bea" };
    var data_label9_3 = xlsxwriter.lxw_chart_data_label{ .value = "Eva" };
    var data_label9_4 = xlsxwriter.lxw_chart_data_label{ .value = "Fay" };
    var data_label9_5 = xlsxwriter.lxw_chart_data_label{ .value = "Liv" };
    var data_label9_6 = xlsxwriter.lxw_chart_data_label{ .value = "Una", .fill = &fill3 };

    // Set the default formatting for the data labels in the series.
    _ = xlsxwriter.chart_series_set_labels_line(series9, &line2);
    _ = xlsxwriter.chart_series_set_labels_fill(series9, &fill2);

    // Create an array of label pointers.
    var data_labels9: [7]?*xlsxwriter.lxw_chart_data_label = undefined;
    data_labels9[0] = &data_label9_1;
    data_labels9[1] = &data_label9_2;
    data_labels9[2] = &data_label9_3;
    data_labels9[3] = &data_label9_4;
    data_labels9[4] = &data_label9_5;
    data_labels9[5] = &data_label9_6;
    data_labels9[6] = null;

    // Set the custom labels.
    _ = xlsxwriter.chart_series_set_labels_custom(series9, &data_labels9);

    // Turn off the legend.
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_NONE);

    // Insert the chart into the worksheet.
    _ = xlsxwriter.worksheet_insert_chart_opt(worksheet, 129, 3, chart, &options);

    // Close the workbook, save the file and free any allocated memory.
    _ = xlsxwriter.workbook_close(workbook);
}
