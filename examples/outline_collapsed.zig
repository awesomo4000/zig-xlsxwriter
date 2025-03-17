//
// Example of how use libxlsxwriter to generate Excel outlines and grouping.
//
// These examples focus mainly on collapsed outlines. See also the outlines.c
// example program for more general examples.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 1-4.
fn createRowExampleData(worksheet: *xlsxwriter.lxw_worksheet, bold: *xlsxwriter.lxw_format) void {
    // Set the column width for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Add data and formulas to the worksheet.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Region", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 5, 0, "North Total", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Sales", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 1000, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 1, 900, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 5, 1, "=SUBTOTAL(9,B2:B5)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet, 6, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 7, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 8, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 9, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 10, 0, "South Total", bold);

    _ = xlsxwriter.worksheet_write_number(worksheet, 6, 1, 400, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 7, 1, 600, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 8, 1, 500, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 9, 1, 600, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 10, 1, "=SUBTOTAL(9,B7:B10)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet, 11, 0, "Grand Total", bold);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 11, 1, "=SUBTOTAL(9,B2:B10)", bold);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 5-6.
fn createColExampleData(worksheet: *xlsxwriter.lxw_worksheet, bold: *xlsxwriter.lxw_format) void {
    // Add data and formulas to the worksheet.
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Month", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 1, "Jan", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 2, "Feb", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 3, "Mar", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 4, "Apr", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 5, "May", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 6, "Jun", null);
    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 7, "Total", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 1, 0, "North", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 1, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 2, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 3, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 4, 25, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 5, 65, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 6, 80, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 1, 7, "=SUM(B2:G2)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 2, 0, "South", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 1, 10, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 2, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 3, 30, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 4, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 5, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 2, 6, 50, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 2, 7, "=SUM(B3:G3)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 3, 0, "East", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 1, 45, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 2, 75, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 3, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 4, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 5, 75, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 3, 6, 100, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 3, 7, "=SUM(B4:G4)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 4, 0, "West", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 1, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 2, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 3, 55, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 4, 35, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 5, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 4, 6, 50, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet, 4, 7, "=SUM(B5:G5)", null);

    _ = xlsxwriter.worksheet_write_formula(worksheet, 5, 7, "=SUM(H2:H5)", bold);
}

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-outline_collapsed.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, "Outlined Rows");
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, "Collapsed Rows 1");
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, "Collapsed Rows 2");
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, "Collapsed Rows 3");
    const worksheet5 = xlsxwriter.workbook_add_worksheet(workbook, "Outline Columns");
    const worksheet6 = xlsxwriter.workbook_add_worksheet(workbook, "Collapsed Columns");

    const bold = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    // Example 1: Create a worksheet with outlined rows. It also includes
    // SUBTOTAL() functions so that it looks like the type of automatic
    // outlines that are generated when you use the 'Sub Totals' option.
    //
    // For outlines the important parameters are 'hidden' and 'level'. Rows
    // with the same 'level' are grouped together. The group will be collapsed
    // if 'hidden' is non-zero.

    // The option structs with the outline level set.
    var options1 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 2,
        .collapsed = 0,
    };
    var options2 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 1,
        .collapsed = 0,
    };

    // Set the row outline properties.
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 1, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 2, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 3, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 4, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 5, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options2);

    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 6, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 7, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 8, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 9, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options1);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet1, 10, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options2);

    // Write the sub-total data that is common to the row examples.
    createRowExampleData(worksheet1, bold);

    // Example 2: Create a worksheet with collapsed outlined rows.
    // This is the same as the example 1 except that all rows are collapsed.
    var options3 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 2,
        .collapsed = 0,
    };
    var options4 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 1,
        .collapsed = 0,
    };
    var options5 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 0,
        .collapsed = 1,
    };

    // Set the row options with the outline level.
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 1, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 2, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 3, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 4, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 5, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options4);

    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 6, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 7, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 8, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 9, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options3);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 10, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options4);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet2, 11, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options5);

    // Write the sub-total data that is common to the row examples.
    createRowExampleData(worksheet2, bold);

    // Example 3: Create a worksheet with collapsed outlined rows. Same as the
    // example 1 except that the two sub-totals are collapsed.
    var options6 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 2,
        .collapsed = 0,
    };
    var options7 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 1,
        .collapsed = 1,
    };

    // Set the row options with the outline level.
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 1, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 2, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 3, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 4, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 5, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options7);

    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 6, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 7, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 8, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 9, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options6);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet3, 10, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options7);

    // Write the sub-total data that is common to the row examples.
    createRowExampleData(worksheet3, bold);

    // Example 4: Create a worksheet with outlined rows. Same as the example 1
    // except that the two sub-totals are collapsed.
    var options8 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 2,
        .collapsed = 0,
    };
    var options9 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 1,
        .collapsed = 1,
    };
    var options10 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 0,
        .collapsed = 1,
    };

    // Set the row options with the outline level.
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 1, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 2, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 3, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 4, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 5, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options9);

    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 6, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 7, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 8, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 9, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options8);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 10, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options9);
    _ = xlsxwriter.worksheet_set_row_opt(worksheet4, 11, xlsxwriter.LXW_DEF_ROW_HEIGHT, null, &options10);

    // Write the sub-total data that is common to the row examples.
    createRowExampleData(worksheet4, bold);

    // Example 5: Create a worksheet with outlined columns.
    var options11 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 1,
        .collapsed = 0,
    };

    // Write the sub-total data that is common to the column examples.
    createColExampleData(worksheet5, bold);

    // Add bold format to the first row.
    _ = xlsxwriter.worksheet_set_row(worksheet5, 0, xlsxwriter.LXW_DEF_ROW_HEIGHT, bold);

    // Set column formatting and the outline level.
    _ = xlsxwriter.worksheet_set_column(worksheet5, 0, 0, 10, bold);
    _ = xlsxwriter.worksheet_set_column_opt(worksheet5, 1, 6, 5, null, &options11);
    _ = xlsxwriter.worksheet_set_column(worksheet5, 7, 7, 10, null);

    // Example 6: Create a worksheet with outlined columns.
    var options12 = xlsxwriter.lxw_row_col_options{
        .hidden = 1,
        .level = 1,
        .collapsed = 0,
    };
    var options13 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 0,
        .collapsed = 1,
    };

    // Write the sub-total data that is common to the column examples.
    createColExampleData(worksheet6, bold);

    // Add bold format to the first row.
    _ = xlsxwriter.worksheet_set_row(worksheet6, 0, xlsxwriter.LXW_DEF_ROW_HEIGHT, bold);

    // Set column formatting and the outline level.
    _ = xlsxwriter.worksheet_set_column(worksheet6, 0, 0, 10, bold);
    _ = xlsxwriter.worksheet_set_column_opt(worksheet6, 1, 6, 5, null, &options12);
    _ = xlsxwriter.worksheet_set_column_opt(worksheet6, 7, 7, 10, null, &options13);

    _ = xlsxwriter.workbook_close(workbook);
}
