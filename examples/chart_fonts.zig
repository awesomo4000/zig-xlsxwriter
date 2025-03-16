//
// An example of a simple Excel chart with user defined fonts using the
// libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_fonts.xlsx");
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
    _ = xlsxwriter.chart_add_series(chart, null, "Sheet1!$A$1:$A$6");

    // Create some fonts to use in the chart
    var font1 = xlsxwriter.lxw_chart_font{
        .name = "Calibri",
        .color = xlsxwriter.LXW_COLOR_BLUE,
        .bold = xlsxwriter.LXW_EXPLICIT_FALSE,
        .italic = xlsxwriter.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font2 = xlsxwriter.lxw_chart_font{
        .name = "Courier",
        .color = 0x92D050,
        .bold = xlsxwriter.LXW_EXPLICIT_FALSE,
        .italic = xlsxwriter.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font3 = xlsxwriter.lxw_chart_font{
        .name = "Arial",
        .color = 0x00B0F0,
        .bold = xlsxwriter.LXW_EXPLICIT_FALSE,
        .italic = xlsxwriter.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font4 = xlsxwriter.lxw_chart_font{
        .name = "Century",
        .color = xlsxwriter.LXW_COLOR_RED,
        .bold = xlsxwriter.LXW_EXPLICIT_FALSE,
        .italic = xlsxwriter.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font5 = xlsxwriter.lxw_chart_font{
        .name = null,
        .color = 0,
        .bold = xlsxwriter.LXW_EXPLICIT_FALSE,
        .italic = xlsxwriter.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = -30,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font6 = xlsxwriter.lxw_chart_font{
        .name = null,
        .color = 0x7030A0,
        .bold = xlsxwriter.LXW_TRUE,
        .italic = xlsxwriter.LXW_TRUE,
        .underline = xlsxwriter.LXW_TRUE,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    // Write the chart title with a font
    _ = xlsxwriter.chart_title_set_name(chart, "Test Results");
    _ = xlsxwriter.chart_title_set_name_font(chart, &font1);

    // Write the Y axis with a font
    _ = xlsxwriter.chart_axis_set_name(chart.*.y_axis, "Units");
    _ = xlsxwriter.chart_axis_set_name_font(chart.*.y_axis, &font2);
    _ = xlsxwriter.chart_axis_set_num_font(chart.*.y_axis, &font3);

    // Write the X axis with a font
    _ = xlsxwriter.chart_axis_set_name(chart.*.x_axis, "Month");
    _ = xlsxwriter.chart_axis_set_name_font(chart.*.x_axis, &font4);
    _ = xlsxwriter.chart_axis_set_num_font(chart.*.x_axis, &font5);

    // Display the chart legend at the bottom of the chart
    _ = xlsxwriter.chart_legend_set_position(chart, xlsxwriter.LXW_CHART_LEGEND_BOTTOM);
    _ = xlsxwriter.chart_legend_set_font(chart, &font6);

    // Insert the chart into the worksheet
    _ = xlsxwriter.worksheet_insert_chart(worksheet, 0, 2, chart);

    _ = xlsxwriter.workbook_close(workbook);
}
