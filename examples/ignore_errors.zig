//
// An example of turning off worksheet cells errors/warnings using
// libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-ignore_errors.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Write strings that looks like numbers. This will cause an Excel warning.
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 2, "123", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 2, "123", null);

    // Write a divide by zero formula. This will also cause an Excel warning.
    _ = xlsxwriter.worksheet_write_formula(worksheet, 4, 2, "=1/0", null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 5, 2, "=1/0", null);

    // Turn off some of the warnings:
    _ = xlsxwriter.worksheet_ignore_errors(worksheet, xlsxwriter.LXW_IGNORE_NUMBER_STORED_AS_TEXT, "C3");
    _ = xlsxwriter.worksheet_ignore_errors(worksheet, xlsxwriter.LXW_IGNORE_EVAL_ERROR, "C6");

    // Write some descriptions for the cells and make the column wider for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet, 1, 1, 16, null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 1, "Warning:", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 1, "Warning turned off:", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 1, "Warning:", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 5, 1, "Warning turned off:", null);

    _ = xlsxwriter.workbook_close(workbook);
}
