//
// Examples of how to add data validation and dropdown lists using the
// libxlsxwriter library.
//
// Data validation is a feature of Excel which allows you to restrict the data
// that a user enters in a cell and to display help and warning messages. It
// also allows you to restrict input to values in a dropdown list.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Write some data to the worksheet.
fn write_worksheet_data(worksheet: *xlsxwriter.lxw_worksheet, format: *xlsxwriter.lxw_format) void {
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Some examples of data validation in libxlsxwriter", format);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Enter values in this column", format);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 3, "Sample Data", format);

    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 3, "Integers", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 4, 1, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 5, 10, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 3, "List data", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 4, "open", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 5, "high", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 6, "close", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 3, "Formula", null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 4, 4, "=AND(F5=50,G5=60)", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 5, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 6, 60, null);
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-data_validate.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Allocate memory for the data validation structure
    var data_validation = std.heap.c_allocator.create(xlsxwriter.lxw_data_validation) catch unreachable;
    // Initialize with zeros
    @memset(@as([*]u8, @ptrCast(data_validation))[0..@sizeOf(xlsxwriter.lxw_data_validation)], 0);

    // Add a format to use to highlight the header cells.
    const format = xlsxwriter.workbook_add_format(workbook);
    xlsxwriter.format_set_border(format, xlsxwriter.LXW_BORDER_THIN);
    xlsxwriter.format_set_fg_color(format, 0xC6EFCE);
    xlsxwriter.format_set_bold(format);
    xlsxwriter.format_set_text_wrap(format);
    xlsxwriter.format_set_align(format, xlsxwriter.LXW_ALIGN_VERTICAL_CENTER);
    xlsxwriter.format_set_indent(format, 1);

    // Write some data for the validations.
    write_worksheet_data(worksheet, format);

    // Set up layout of the worksheet.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 55, null);
    _ = xlsxwriter.worksheet_set_column(worksheet, 1, 1, 15, null);
    _ = xlsxwriter.worksheet_set_column(worksheet, 3, 3, 15, null);
    _ = xlsxwriter.worksheet_set_row(worksheet, 0, 36, null);

    // Example 1. Limiting input to an integer in a fixed range.
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "Enter an integer between 1 and 10", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 10;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 2, 1, data_validation);

    // Example 2. Limiting input to an integer outside a fixed range.
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 0, "Enter an integer that is not between 1 and 10 (using cell references)", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER_FORMULA;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_NOT_BETWEEN;
    data_validation.minimum_formula = "=E3";
    data_validation.maximum_formula = "=F3";

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 4, 1, data_validation);

    // Example 3. Limiting input to an integer greater than a fixed value.
    _ = xlsxwriter.worksheet_write_string(worksheet, 6, 0, "Enter an integer greater than 0", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation.value_number = 0;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 6, 1, data_validation);

    // Example 4. Limiting input to an integer less than a fixed value.
    _ = xlsxwriter.worksheet_write_string(worksheet, 8, 0, "Enter an integer less than 10", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_LESS_THAN;
    data_validation.value_number = 10;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 8, 1, data_validation);

    // Example 5. Limiting input to a decimal in a fixed range.
    _ = xlsxwriter.worksheet_write_string(worksheet, 10, 0, "Enter a decimal between 0.1 and 0.5", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_DECIMAL;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_number = 0.1;
    data_validation.maximum_number = 0.5;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 10, 1, data_validation);

    // Example 6. Limiting input to a value in a dropdown list.
    _ = xlsxwriter.worksheet_write_string(worksheet, 12, 0, "Select a value from a dropdown list", null);

    const list = [_][*c]const u8{ "open", "high", "close", null };
    const list_ptr: [*c][*c]const u8 = @ptrCast(@constCast(&list));

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_LIST;
    data_validation.value_list = list_ptr;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 12, 1, data_validation);

    // Example 7. Limiting input to a value in a dropdown list.
    _ = xlsxwriter.worksheet_write_string(worksheet, 14, 0, "Select a value from a dropdown list (using a cell range)", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_LIST_FORMULA;
    data_validation.value_formula = "=$E$4:$G$4";

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 14, 1, data_validation);

    // Example 8. Limiting input to a date in a fixed range.
    _ = xlsxwriter.worksheet_write_string(worksheet, 16, 0, "Enter a date between 1/1/2024 and 12/12/2024", null);

    const datetime1 = xlsxwriter.lxw_datetime{ .year = 2024, .month = 1, .day = 1, .hour = 0, .min = 0, .sec = 0 };
    const datetime2 = xlsxwriter.lxw_datetime{ .year = 2024, .month = 12, .day = 12, .hour = 0, .min = 0, .sec = 0 };

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_DATE;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_datetime = datetime1;
    data_validation.maximum_datetime = datetime2;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 16, 1, data_validation);

    // Example 9. Limiting input to a time in a fixed range.
    _ = xlsxwriter.worksheet_write_string(worksheet, 18, 0, "Enter a time between 6:00 and 12:00", null);

    const datetime3 = xlsxwriter.lxw_datetime{ .year = 0, .month = 0, .day = 0, .hour = 6, .min = 0, .sec = 0 };
    const datetime4 = xlsxwriter.lxw_datetime{ .year = 0, .month = 0, .day = 0, .hour = 12, .min = 0, .sec = 0 };

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_TIME;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_datetime = datetime3;
    data_validation.maximum_datetime = datetime4;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 18, 1, data_validation);

    // Example 10. Limiting input to a string greater than a fixed length.
    _ = xlsxwriter.worksheet_write_string(worksheet, 20, 0, "Enter a string longer than 3 characters", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_LENGTH;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation.value_number = 3;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 20, 1, data_validation);

    // Example 11. Limiting input based on a formula.
    _ = xlsxwriter.worksheet_write_string(worksheet, 22, 0, "Enter a value if the following is true \"=AND(F5=50,G5=60)\"", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_CUSTOM_FORMULA;
    data_validation.value_formula = "=AND(F5=50,G5=60)";

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 22, 1, data_validation);

    // Example 12. Displaying and modifying data validation messages.
    _ = xlsxwriter.worksheet_write_string(worksheet, 24, 0, "Displays a message when you select the cell", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 24, 1, data_validation);

    // Example 13. Displaying and modifying data validation messages.
    _ = xlsxwriter.worksheet_write_string(worksheet, 26, 0, "Display a custom error message when integer isn't between 1 and 100", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 26, 1, data_validation);

    // Example 14. Displaying and modifying data validation messages.
    _ = xlsxwriter.worksheet_write_string(worksheet, 28, 0, "Display a custom info message when integer isn't between 1 and 100", null);

    data_validation.validate = xlsxwriter.LXW_VALIDATION_TYPE_INTEGER;
    data_validation.criteria = xlsxwriter.LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";
    data_validation.error_type = xlsxwriter.LXW_VALIDATION_ERROR_TYPE_INFORMATION;

    _ = xlsxwriter.worksheet_data_validation_cell(worksheet, 28, 1, data_validation);

    // Cleanup.
    std.heap.c_allocator.destroy(data_validation);

    _ = xlsxwriter.workbook_close(workbook);
}
