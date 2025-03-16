//
// Example of adding an autofilter to a worksheet in Excel using
// libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const c = @cImport({
    @cInclude("string.h");
});
const xlsxwriter = @import("xlsxwriter");

fn writeWorksheetHeader(worksheet: ?*xlsxwriter.lxw_worksheet, header: ?*xlsxwriter.lxw_format) void {
    // Make the columns wider for clarity
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 3, 12, null);

    // Write the column headers
    _ = xlsxwriter.worksheet_set_row(worksheet, 0, 20, header);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Region", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Item", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "Volume", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 3, "Month", null);
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-autofilter.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet5 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet6 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet7 = xlsxwriter.workbook_add_worksheet(workbook, null);

    const Row = struct {
        region: [*:0]const u8,
        item: [*:0]const u8,
        volume: i32,
        month: [*:0]const u8,
    };

    const data = [_]Row{
        .{ .region = "East", .item = "Apple", .volume = 9000, .month = "July" },
        .{ .region = "East", .item = "Apple", .volume = 5000, .month = "July" },
        .{ .region = "South", .item = "Orange", .volume = 9000, .month = "September" },
        .{ .region = "North", .item = "Apple", .volume = 2000, .month = "November" },
        .{ .region = "West", .item = "Apple", .volume = 9000, .month = "November" },
        .{ .region = "South", .item = "Pear", .volume = 7000, .month = "October" },
        .{ .region = "North", .item = "Pear", .volume = 9000, .month = "August" },
        .{ .region = "West", .item = "Orange", .volume = 1000, .month = "December" },
        .{ .region = "West", .item = "Grape", .volume = 1000, .month = "November" },
        .{ .region = "South", .item = "Pear", .volume = 10000, .month = "April" },
        .{ .region = "West", .item = "Grape", .volume = 6000, .month = "January" },
        .{ .region = "South", .item = "Orange", .volume = 3000, .month = "May" },
        .{ .region = "North", .item = "Apple", .volume = 3000, .month = "December" },
        .{ .region = "South", .item = "Apple", .volume = 7000, .month = "February" },
        .{ .region = "West", .item = "Grape", .volume = 1000, .month = "December" },
        .{ .region = "East", .item = "Grape", .volume = 8000, .month = "February" },
        .{ .region = "South", .item = "Grape", .volume = 10000, .month = "June" },
        .{ .region = "West", .item = "Pear", .volume = 7000, .month = "December" },
        .{ .region = "South", .item = "Apple", .volume = 2000, .month = "October" },
        .{ .region = "East", .item = "Grape", .volume = 7000, .month = "December" },
        .{ .region = "North", .item = "Grape", .volume = 6000, .month = "April" },
        .{ .region = "East", .item = "Pear", .volume = 8000, .month = "February" },
        .{ .region = "North", .item = "Apple", .volume = 7000, .month = "August" },
        .{ .region = "North", .item = "Orange", .volume = 7000, .month = "July" },
        .{ .region = "North", .item = "Apple", .volume = 6000, .month = "June" },
        .{ .region = "South", .item = "Grape", .volume = 8000, .month = "September" },
        .{ .region = "West", .item = "Apple", .volume = 3000, .month = "October" },
        .{ .region = "South", .item = "Orange", .volume = 10000, .month = "November" },
        .{ .region = "West", .item = "Grape", .volume = 4000, .month = "July" },
        .{ .region = "North", .item = "Orange", .volume = 5000, .month = "August" },
        .{ .region = "East", .item = "Orange", .volume = 1000, .month = "November" },
        .{ .region = "East", .item = "Orange", .volume = 4000, .month = "October" },
        .{ .region = "North", .item = "Grape", .volume = 5000, .month = "August" },
        .{ .region = "East", .item = "Apple", .volume = 1000, .month = "December" },
        .{ .region = "South", .item = "Apple", .volume = 10000, .month = "March" },
        .{ .region = "East", .item = "Grape", .volume = 7000, .month = "October" },
        .{ .region = "West", .item = "Grape", .volume = 1000, .month = "September" },
        .{ .region = "East", .item = "Grape", .volume = 10000, .month = "October" },
        .{ .region = "South", .item = "Orange", .volume = 8000, .month = "March" },
        .{ .region = "North", .item = "Apple", .volume = 4000, .month = "July" },
        .{ .region = "South", .item = "Orange", .volume = 5000, .month = "July" },
        .{ .region = "West", .item = "Apple", .volume = 4000, .month = "June" },
        .{ .region = "East", .item = "Apple", .volume = 5000, .month = "April" },
        .{ .region = "North", .item = "Pear", .volume = 3000, .month = "August" },
        .{ .region = "East", .item = "Grape", .volume = 9000, .month = "November" },
        .{ .region = "North", .item = "Orange", .volume = 8000, .month = "October" },
        .{ .region = "East", .item = "Apple", .volume = 10000, .month = "June" },
        .{ .region = "South", .item = "Pear", .volume = 1000, .month = "December" },
        .{ .region = "North", .item = "Grape", .volume = 10000, .month = "July" },
        .{ .region = "East", .item = "Grape", .volume = 6000, .month = "February" },
    };

    var hidden = xlsxwriter.lxw_row_col_options{ .hidden = xlsxwriter.LXW_TRUE };

    const header = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(header);

    //
    // Example 1. Autofilter without conditions.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet1, header);

    // Write the row data
    for (data, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet1, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet1, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet1, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet1, @intCast(i + 1), 3, row.month, null);
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet1, 0, 0, 50, 3);

    //
    // Example 2. Autofilter with a filter condition in the first column.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet2, header);

    // Write the row data
    for (data, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet2, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet2, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet2, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet2, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (c.strcmp(row.region, "East") != 0) {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet2, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet2, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule2 = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_EQUAL_TO,
        .value_string = "East",
        .value = 0,
    };

    _ = xlsxwriter.worksheet_filter_column(worksheet2, 0, &filter_rule2);

    //
    // Example 3. Autofilter with a dual filter condition in one of the columns.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet3, header);

    // Write the row data
    for (data, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet3, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet3, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet3, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet3, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (c.strcmp(row.region, "East") != 0 and c.strcmp(row.region, "South") != 0) {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet3, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet3, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule3a = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_EQUAL_TO,
        .value_string = "East",
        .value = 0,
    };

    var filter_rule3b = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_EQUAL_TO,
        .value_string = "South",
        .value = 0,
    };

    _ = xlsxwriter.worksheet_filter_column2(worksheet3, 0, &filter_rule3a, &filter_rule3b, xlsxwriter.LXW_FILTER_OR);

    //
    // Example 4. Autofilter with filter conditions in two columns.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet4, header);

    // Write the row data
    for (data, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet4, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet4, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet4, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet4, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (!(c.strcmp(row.region, "East") == 0 and
            row.volume > 3000 and row.volume < 8000))
        {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet4, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet4, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule4a = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_EQUAL_TO,
        .value_string = "East",
        .value = 0,
    };

    var filter_rule4b = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_GREATER_THAN,
        .value_string = "",
        .value = 3000,
    };

    var filter_rule4c = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_LESS_THAN,
        .value_string = "",
        .value = 8000,
    };

    _ = xlsxwriter.worksheet_filter_column(worksheet4, 0, &filter_rule4a);
    _ = xlsxwriter.worksheet_filter_column2(worksheet4, 2, &filter_rule4b, &filter_rule4c, xlsxwriter.LXW_FILTER_AND);

    //
    // Example 5. Autofilter with a list filter condition.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet5, header);

    // Write the row data
    for (data, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet5, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet5, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet5, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet5, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (c.strcmp(row.region, "East") != 0 and
            c.strcmp(row.region, "North") != 0 and
            c.strcmp(row.region, "South") != 0)
        {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet5, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet5, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule5 = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_EQUAL_TO,
        .value_string = "East",
        .value = 0,
    };

    _ = xlsxwriter.worksheet_filter_column(worksheet5, 0, &filter_rule5);

    //
    // Example 6. Autofilter with filter for blanks.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet6, header);

    // Create a copy of data with one blank region
    var data_with_blank = data;
    var blank_row = data_with_blank[5];
    blank_row.region = "";
    data_with_blank[5] = blank_row;

    // Write the row data
    for (data_with_blank, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet6, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet6, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet6, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet6, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (c.strcmp(row.region, "") != 0) {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet6, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet6, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule6 = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_BLANKS,
        .value_string = "",
        .value = 0,
    };

    _ = xlsxwriter.worksheet_filter_column(worksheet6, 0, &filter_rule6);

    //
    // Example 7. Autofilter with filter for non-blanks.
    //

    // Set up the worksheet data
    writeWorksheetHeader(worksheet7, header);

    // Write the row data
    for (data_with_blank, 0..) |row, i| {
        _ = xlsxwriter.worksheet_write_string(worksheet7, @intCast(i + 1), 0, row.region, null);
        _ = xlsxwriter.worksheet_write_string(worksheet7, @intCast(i + 1), 1, row.item, null);
        _ = xlsxwriter.worksheet_write_number(worksheet7, @intCast(i + 1), 2, @floatFromInt(row.volume), null);
        _ = xlsxwriter.worksheet_write_string(worksheet7, @intCast(i + 1), 3, row.month, null);

        // Hide rows that don't match the filter
        if (c.strcmp(row.region, "") == 0) {
            _ = xlsxwriter.worksheet_set_row_opt(worksheet7, @intCast(i + 1), xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &hidden);
        }
    }

    // Add the autofilter
    _ = xlsxwriter.worksheet_autofilter(worksheet7, 0, 0, 50, 3);

    // Add the filter criteria
    var filter_rule7 = xlsxwriter.lxw_filter_rule{
        .criteria = xlsxwriter.LXW_FILTER_CRITERIA_NON_BLANKS,
        .value_string = "",
        .value = 0,
    };

    _ = xlsxwriter.worksheet_filter_column(worksheet7, 0, &filter_rule7);

    _ = xlsxwriter.workbook_close(workbook);
}
