//
// An example showing all 48 default chart styles available in Excel 2007
// using the libxlsxwriter library. Note, these styles are not the same as the
// styles available in Excel 2013.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-chart_styles.xlsx");

    // Define chart types and names
    const chart_types = [_]u8{
        xlsxwriter.LXW_CHART_COLUMN,
        xlsxwriter.LXW_CHART_AREA,
        xlsxwriter.LXW_CHART_LINE,
        xlsxwriter.LXW_CHART_PIE,
    };
    const chart_names = [_][:0]const u8{ "Column", "Area", "Line", "Pie" };

    // Create a worksheet for each chart type
    for (chart_types, chart_names) |chart_type, chart_name| {
        const worksheet = xlsxwriter.workbook_add_worksheet(workbook, chart_name);
        _ = xlsxwriter.worksheet_set_zoom(worksheet, 30);

        // Create 48 charts, each with a different style
        var style_num: u8 = 1;
        var row_num: usize = 0;
        while (row_num < 90) : (row_num += 15) {
            var col_num: usize = 0;
            while (col_num < 64) : (col_num += 8) {
                const chart = xlsxwriter.workbook_add_chart(workbook, chart_type);

                // Create chart title with style number
                var title_buf: [32]u8 = undefined;
                const title = std.fmt.bufPrintZ(&title_buf, "Style {d}", .{style_num}) catch unreachable;

                _ = xlsxwriter.chart_add_series(chart, null, "=Data!$A$1:$A$6");
                _ = xlsxwriter.chart_title_set_name(chart, title.ptr);
                _ = xlsxwriter.chart_set_style(chart, style_num);

                _ = xlsxwriter.worksheet_insert_chart(worksheet, @intCast(row_num), @intCast(col_num), chart);

                style_num += 1;
            }
        }
    }

    // Create a worksheet with data for the charts
    const data_worksheet = xlsxwriter.workbook_add_worksheet(workbook, "Data");
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 0, 0, 10, null);
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 1, 0, 40, null);
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 2, 0, 50, null);
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 3, 0, 20, null);
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 4, 0, 10, null);
    _ = xlsxwriter.worksheet_write_number(data_worksheet, 5, 0, 50, null);

    _ = xlsxwriter.workbook_close(workbook);
}
