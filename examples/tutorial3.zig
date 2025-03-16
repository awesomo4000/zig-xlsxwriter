const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Some data we want to write to the worksheet.
const Expense = struct {
    item: []const u8,
    cost: i32,
    datetime: xlsxwriter.lxw_datetime,
};

pub fn main() !void {
    // Create a workbook and add a worksheet.
    const workbook = xlsxwriter.workbook_new("zig-tutorial3.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);
    var row: u32 = 0;
    const col: u32 = 0;

    // Add a bold format to use to highlight cells.
    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Add a number format for cells with money.
    const money = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(money, "$#,##0");

    // Add an Excel date format.
    const date_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_num_format(date_format, "mmmm d yyyy");

    // Adjust the column width.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 15, null);

    // Write some data header.
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col, "Item", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 1, "Cost", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, row, col + 2, "Date", bold);

    // Create our expense data
    const expenses = [_]Expense{
        .{ .item = "Rent", .cost = 1000, .datetime = .{ .year = 2013, .month = 1, .day = 13 } },
        .{ .item = "Gas", .cost = 100, .datetime = .{ .year = 2013, .month = 1, .day = 14 } },
        .{ .item = "Food", .cost = 300, .datetime = .{ .year = 2013, .month = 1, .day = 16 } },
        .{ .item = "Gym", .cost = 50, .datetime = .{ .year = 2013, .month = 1, .day = 20 } },
    };

    // Iterate over the data and write it out element by element.
    for (expenses, 0..) |expense, i| {
        // Write from the first cell below the headers.
        row = @intCast(i + 1);
        _ = xlsxwriter.worksheet_write_string(worksheet, row, col, expense.item.ptr, null);

        // Create a mutable copy of the datetime for the C function
        var datetime_copy = expense.datetime;
        _ = xlsxwriter.worksheet_write_datetime(worksheet, row, col + 1, &datetime_copy, date_format);

        _ = xlsxwriter.worksheet_write_number(worksheet, row, col + 2, @floatFromInt(expense.cost), money);
    }

    // Write a total using a formula.
    _ = xlsxwriter.worksheet_write_string(worksheet, row + 1, col, "Total", bold);
    _ = xlsxwriter.worksheet_write_formula(worksheet, row + 1, col + 2, "=SUM(C2:C5)", money);

    // Save the workbook and free any allocated memory.
    _ = xlsxwriter.workbook_close(workbook);
}
