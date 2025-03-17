//
// An example of how to add conditional formatting to an libxlsxwriter file.
//
// Conditional formatting allows you to apply a format to a cell or a
// range of cells based on certain criteria.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Write some data to the worksheet.
fn writeWorksheetData(worksheet: *xlsxwriter.lxw_worksheet) void {
    const data = [10][10]u8{
        [_]u8{ 34, 72, 38, 30, 75, 48, 75, 66, 84, 86 },
        [_]u8{ 6, 24, 1, 84, 54, 62, 60, 3, 26, 59 },
        [_]u8{ 28, 79, 97, 13, 85, 93, 93, 22, 5, 14 },
        [_]u8{ 27, 71, 40, 17, 18, 79, 90, 93, 29, 47 },
        [_]u8{ 88, 25, 33, 23, 67, 1, 59, 79, 47, 36 },
        [_]u8{ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 },
        [_]u8{ 6, 57, 88, 28, 10, 26, 37, 7, 41, 48 },
        [_]u8{ 52, 78, 1, 96, 26, 45, 47, 33, 96, 36 },
        [_]u8{ 60, 54, 81, 66, 81, 90, 80, 93, 12, 55 },
        [_]u8{ 70, 5, 46, 14, 71, 19, 66, 36, 41, 21 },
    };

    for (0..10) |row| {
        for (0..10) |col| {
            _ = xlsxwriter.worksheet_write_number(worksheet, @intCast(row + 2), @intCast(col + 1), @floatFromInt(data[row][col]), null);
        }
    }
}

// Reset the conditional format options back to their initial state.
fn resetConditionalFormat(conditional_format: *xlsxwriter.lxw_conditional_format) void {
    @memset(@as([*]u8, @ptrCast(conditional_format))[0..@sizeOf(xlsxwriter.lxw_conditional_format)], 0);
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-conditional_format2.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet5 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet6 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet7 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet8 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet9 = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Add a format. Light red fill with dark red text.
    const format1 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bg_color(format1, 0xFFC7CE);
    _ = xlsxwriter.format_set_font_color(format1, 0x9C0006);

    // Add a format. Green fill with dark green text.
    const format2 = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bg_color(format2, 0xC6EFCE);
    _ = xlsxwriter.format_set_font_color(format2, 0x006100);

    // Create a single conditional format object to reuse in the examples.
    var conditional_format: xlsxwriter.lxw_conditional_format = undefined;
    resetConditionalFormat(&conditional_format);

    // Example 1. Conditional formatting based on simple cell based criteria.
    writeWorksheetData(worksheet1);

    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        0,
        0,
        "Cells with values >= 50 are in light red. Values < 50 are in light green.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
    conditional_format.value = 50;
    conditional_format.format = format1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet1, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_LESS_THAN;
    conditional_format.value = 50;
    conditional_format.format = format2;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet1, 2, 1, 11, 10, &conditional_format);

    // Example 2. Conditional formatting based on max and min values.
    writeWorksheetData(worksheet2);

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        0,
        0,
        "Values between 30 and 70 are in light red. Values outside that range are in light green.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_BETWEEN;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format = format1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet2, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format = format2;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet2, 2, 1, 11, 10, &conditional_format);

    // Example 3. Conditional formatting with duplicate and unique values.
    writeWorksheetData(worksheet3);

    _ = xlsxwriter.worksheet_write_string(
        worksheet3,
        0,
        0,
        "Duplicate values are in light red. Unique values are in light green.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_DUPLICATE;
    conditional_format.format = format1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet3, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_UNIQUE;
    conditional_format.format = format2;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet3, 2, 1, 11, 10, &conditional_format);

    // Example 4. Conditional formatting with above and below average values.
    writeWorksheetData(worksheet4);

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        0,
        0,
        "Above average values are in light red. Below average values are in light green.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_AVERAGE;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE;
    conditional_format.format = format1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet4, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_AVERAGE;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW;
    conditional_format.format = format2;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet4, 2, 1, 11, 10, &conditional_format);

    // Example 5. Conditional formatting with top and bottom values.
    writeWorksheetData(worksheet5);

    _ = xlsxwriter.worksheet_write_string(
        worksheet5,
        0,
        0,
        "Top 10 values are in light red. Bottom 10 values are in light green.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_TOP;
    conditional_format.value = 10;
    conditional_format.format = format1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet5, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_BOTTOM;
    conditional_format.value = 10;
    conditional_format.format = format2;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet5, 2, 1, 11, 10, &conditional_format);

    // Example 6. Conditional formatting with multiple ranges.
    writeWorksheetData(worksheet6);

    _ = xlsxwriter.worksheet_write_string(
        worksheet6,
        0,
        0,
        "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.",
        null,
    );

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
    conditional_format.value = 50;
    conditional_format.format = format1;
    conditional_format.multi_range = "B3:K6 B9:K12";
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet6, 2, 1, 11, 10, &conditional_format);

    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_CELL;
    conditional_format.criteria = xlsxwriter.LXW_CONDITIONAL_CRITERIA_LESS_THAN;
    conditional_format.value = 50;
    conditional_format.format = format2;
    conditional_format.multi_range = "B3:K6 B9:K12";
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet6, 2, 1, 11, 10, &conditional_format);

    // Reset the options before the next example.
    resetConditionalFormat(&conditional_format);

    // Example 7. Conditional formatting with 2 color scales.
    // Write the worksheet data.
    for (1..13) |i| {
        _ = xlsxwriter.worksheet_write_number(worksheet7, @intCast(i + 1), 1, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet7, @intCast(i + 1), 3, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet7, @intCast(i + 1), 6, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet7, @intCast(i + 1), 8, @floatFromInt(i), null);
    }

    _ = xlsxwriter.worksheet_write_string(
        worksheet7,
        0,
        0,
        "Examples of color scales with default and user colors.",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(worksheet7, 1, 1, "2 Color Scale", null);
    _ = xlsxwriter.worksheet_write_string(worksheet7, 1, 3, "2 Color Scale + user colors", null);
    _ = xlsxwriter.worksheet_write_string(worksheet7, 1, 6, "3 Color Scale", null);
    _ = xlsxwriter.worksheet_write_string(worksheet7, 1, 8, "3 Color Scale + user colors", null);

    // 2 color scale with standard colors.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_2_COLOR_SCALE;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet7, 2, 1, 13, 1, &conditional_format);

    // 2 color scale with user defined colors.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_2_COLOR_SCALE;
    conditional_format.min_color = 0xFF0000;
    conditional_format.max_color = 0x00FF00;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet7, 2, 3, 13, 3, &conditional_format);

    // Reset the colors before the next example.
    resetConditionalFormat(&conditional_format);

    // 3 color scale with standard colors.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_3_COLOR_SCALE;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet7, 2, 6, 13, 6, &conditional_format);

    // 3 color scale with user defined colors.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_3_COLOR_SCALE;
    conditional_format.min_color = 0xFF0000;
    conditional_format.mid_color = 0xFFFF00;
    conditional_format.max_color = 0x00FF00;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet7, 2, 8, 13, 8, &conditional_format);

    // Example 8. Conditional formatting with data bars.
    // First data bar example.
    _ = xlsxwriter.worksheet_write_string(
        worksheet8,
        0,
        0,
        "Examples of data bars.",
        null,
    );

    // Write the worksheet data.
    for (1..13) |i| {
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 1, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 3, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 5, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 7, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 9, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 11, @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet8, @intCast(i + 1), 13, @floatFromInt(i), null);
    }

    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 1, "Default data bars", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 3, "Data bars + border", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 5, "Bars with user color", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 7, "Negative same as positive", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 9, "Zero axis", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 11, "Right to left", null);
    _ = xlsxwriter.worksheet_write_string(worksheet8, 1, 13, "Excel 2010 style", null);

    // Default data bars.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 1, 13, 1, &conditional_format);

    // Data bars with border.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.bar_border_color = 0;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 3, 13, 3, &conditional_format);

    // User defined color.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.bar_color = 0x63C384;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 5, 13, 5, &conditional_format);

    // Same color for negative values.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.bar_negative_color_same = 1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 7, 13, 7, &conditional_format);

    // Zero axis.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.bar_axis_position = xlsxwriter.LXW_CONDITIONAL_BAR_AXIS_MIDPOINT;
    conditional_format.bar_axis_color = 0;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 9, 13, 9, &conditional_format);

    // Right to left.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.bar_direction = xlsxwriter.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 11, 13, 11, &conditional_format);

    // Excel 2010 style.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_DATA_BAR;
    conditional_format.data_bar_2010 = 1;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet8, 2, 13, 13, 13, &conditional_format);

    // Example 9. Conditional formatting with icon sets.
    _ = xlsxwriter.worksheet_write_string(
        worksheet9,
        0,
        0,
        "Examples of conditional formats with icon sets.",
        null,
    );

    // Write the worksheet data.
    for (1..4) |i| {
        _ = xlsxwriter.worksheet_write_number(worksheet9, 2, @intCast(i), @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet9, 3, @intCast(i), @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet9, 4, @intCast(i), @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet9, 5, @intCast(i), @floatFromInt(i), null);
    }

    for (1..5) |i| {
        _ = xlsxwriter.worksheet_write_number(worksheet9, 6, @intCast(i), @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet9, 7, @intCast(i), @floatFromInt(i), null);
        _ = xlsxwriter.worksheet_write_number(worksheet9, 8, @intCast(i), @floatFromInt(i), null);
    }

    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 1, "3 Traffic lights", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 3, "3 Traffic lights unrimmed", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 5, "3 Arrows", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 7, "3 Symbols circled", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 9, "3 Symbols", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 11, "3 Flags", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 13, "3 Traffic lights", null);

    // Reset the conditional format.
    resetConditionalFormat(&conditional_format);

    // Three traffic lights (default style).
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 1, 4, 1, &conditional_format);

    // Three traffic lights (unrimmed style).
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 3, 4, 3, &conditional_format);

    // Three arrows.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 5, 4, 5, &conditional_format);

    // Three symbols circled.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 7, 4, 7, &conditional_format);

    // Three symbols.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 9, 4, 9, &conditional_format);

    // Three flags.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_FLAGS;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 11, 4, 11, &conditional_format);

    // Three traffic lights.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 13, 4, 13, &conditional_format);

    // Examples of 4 set icons.
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 15, "4 Traffic lights", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 17, "4 Arrows", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 19, "4 Red-Black", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 1, 21, "4 Ratings", null);

    // Four traffic lights.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 15, 5, 15, &conditional_format);

    // Four arrows.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 17, 5, 17, &conditional_format);

    // Four red to black.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 19, 5, 19, &conditional_format);

    // Four ratings.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_4_RATINGS;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 2, 21, 5, 21, &conditional_format);

    // Examples of 5 set icons.
    _ = xlsxwriter.worksheet_write_string(worksheet9, 6, 15, "5 Arrows", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 6, 17, "5 Ratings", null);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 6, 19, "5 Quarters", null);

    // Five arrows.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 7, 15, 11, 15, &conditional_format);

    // Five ratings.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_5_RATINGS;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 7, 17, 11, 17, &conditional_format);

    // Five quarters.
    conditional_format.type = xlsxwriter.LXW_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format.icon_style = xlsxwriter.LXW_CONDITIONAL_ICONS_5_QUARTERS;
    _ = xlsxwriter.worksheet_conditional_format_range(worksheet9, 7, 19, 11, 19, &conditional_format);

    _ = xlsxwriter.workbook_close(workbook);
}
