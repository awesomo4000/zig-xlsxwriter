//
// An example of using the new Excel LAMBDA() function with the libxlsxwriter
// library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-lambda.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Convert cell references to row/col format
    const row_a1 = xlsxwriter.lxw_name_to_row("A1");
    const col_a1 = xlsxwriter.lxw_name_to_col("A1");
    const row_a2 = xlsxwriter.lxw_name_to_row("A2");
    const col_a2 = xlsxwriter.lxw_name_to_col("A2");

    // Note that the formula name is prefixed with "_xlfn." and that the
    // lambda function parameters are prefixed with "_xlpm.". These prefixes
    // won't show up in Excel.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet,
        row_a1,
        col_a1,
        "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)",
        null,
    );

    // Create the lambda function as a defined name and write it as a dynamic formula
    _ = xlsxwriter.workbook_define_name(
        workbook,
        "ToCelsius",
        "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))",
    );

    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet,
        row_a2,
        col_a2,
        "=ToCelsius(212)",
        null,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
