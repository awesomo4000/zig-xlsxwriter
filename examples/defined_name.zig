//
// Example of how to create defined names using libxlsxwriter. This method is
// used to define a user friendly name to represent a value, a single cell or
// a range of cells in a workbook.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-defined_name.xlsx");

    // Add two worksheets
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Define some global/workbook names
    _ = xlsxwriter.workbook_define_name(workbook, "Exchange_rate", "=0.96");
    _ = xlsxwriter.workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10");

    // Define a local/worksheet name. This overrides the global "Sales" name
    // with a local defined name.
    _ = xlsxwriter.workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10");

    // Write some text to the worksheets and one of the defined names in a formula
    // Process worksheet1
    _ = xlsxwriter.worksheet_set_column(
        worksheet1,
        0,
        0,
        45,
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        0,
        0,
        "This worksheet contains some defined names.",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        1,
        0,
        "See Formulas -> Name Manager above.",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        2,
        0,
        "Example formula in cell B3 ->",
        null,
    );

    _ = xlsxwriter.worksheet_write_formula(
        worksheet1,
        2,
        1,
        "=Exchange_rate",
        null,
    );

    // Process worksheet2
    _ = xlsxwriter.worksheet_set_column(
        worksheet2,
        0,
        0,
        45,
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        0,
        0,
        "This worksheet contains some defined names.",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        1,
        0,
        "See Formulas -> Name Manager above.",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        2,
        0,
        "Example formula in cell B3 ->",
        null,
    );

    _ = xlsxwriter.worksheet_write_formula(
        worksheet2,
        2,
        1,
        "=Exchange_rate",
        null,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
