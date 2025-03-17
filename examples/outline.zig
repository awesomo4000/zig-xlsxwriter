//
// Example of how use libxlsxwriter to generate Excel outlines and grouping.
//
// Excel allows you to group rows or columns so that they can be hidden or
// displayed with a single mouse click. This feature is referred to as
// outlines.
//
// Outlines can reduce complex data down to a few salient sub-totals or
// summaries.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-outline.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, "Outlined Rows");
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, "Collapsed Rows");
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, "Outline Columns");
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, "Outline levels");

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

    // Set the column width for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet1, 0, 0, 20, null);

    // Set the row options with the outline level.
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        1,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        2,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        3,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        4,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        5,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options2,
    );

    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        6,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        7,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        8,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        9,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet1,
        10,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options2,
    );

    // Add data and formulas to the worksheet.
    _ = xlsxwriter.worksheet_write_string(worksheet1, 0, 0, "Region", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 1, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 2, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 3, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 4, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 5, 0, "North Total", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet1, 0, 1, "Sales", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 1, 1, 1000, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 2, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 3, 1, 900, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 4, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet1, 5, 1, "=SUBTOTAL(9,B2:B5)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet1, 6, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 7, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 8, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 9, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 10, 0, "South Total", bold);

    _ = xlsxwriter.worksheet_write_number(worksheet1, 6, 1, 400, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 7, 1, 600, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 8, 1, 500, null);
    _ = xlsxwriter.worksheet_write_number(worksheet1, 9, 1, 600, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet1, 10, 1, "=SUBTOTAL(9,B7:B10)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet1, 11, 0, "Grand Total", bold);
    _ = xlsxwriter.worksheet_write_formula(worksheet1, 11, 1, "=SUBTOTAL(9,B2:B10)", bold);

    // Example 2: Create a worksheet with outlined rows. This is the same as
    // the previous example except that the rows are collapsed. Note: We need
    // to indicate the row that contains the collapsed symbol '+' with the
    // optional parameter, 'collapsed'.

    // The option structs with the outline level and collapsed property set.
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

    // Set the column width for clarity.
    _ = xlsxwriter.worksheet_set_column(worksheet2, 0, 0, 20, null);

    // Set the row options with the outline level.
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        1,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        2,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        3,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        4,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        5,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options4,
    );

    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        6,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        7,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        8,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        9,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        10,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options4,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet2,
        11,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &options5,
    );

    // Add data and formulas to the worksheet.
    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 0, "Region", bold);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 1, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 2, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 3, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 4, 0, "North", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 5, 0, "North Total", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 1, "Sales", bold);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 1, 1, 1000, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 2, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 3, 1, 900, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 4, 1, 1200, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet2, 5, 1, "=SUBTOTAL(9,B2:B5)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet2, 6, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 7, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 8, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 9, 0, "South", null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 10, 0, "South Total", bold);

    _ = xlsxwriter.worksheet_write_number(worksheet2, 6, 1, 400, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 7, 1, 600, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 8, 1, 500, null);
    _ = xlsxwriter.worksheet_write_number(worksheet2, 9, 1, 600, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet2, 10, 1, "=SUBTOTAL(9,B7:B10)", bold);

    _ = xlsxwriter.worksheet_write_string(worksheet2, 11, 0, "Grand Total", bold);
    _ = xlsxwriter.worksheet_write_formula(worksheet2, 11, 1, "=SUBTOTAL(9,B2:B10)", bold);

    // Example 3: Create a worksheet with outlined columns.
    var options6 = xlsxwriter.lxw_row_col_options{
        .hidden = 0,
        .level = 1,
        .collapsed = 0,
    };

    // Add data and formulas to the worksheet.
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 0, "Month", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 1, "Jan", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 2, "Feb", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 3, "Mar", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 4, "Apr", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 5, "May", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 6, "Jun", null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 7, "Total", null);

    _ = xlsxwriter.worksheet_write_string(worksheet3, 1, 0, "North", null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 1, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 2, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 3, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 4, 25, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 5, 65, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 1, 6, 80, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet3, 1, 7, "=SUM(B2:G2)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet3, 2, 0, "South", null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 1, 10, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 2, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 3, 30, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 4, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 5, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 2, 6, 50, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet3, 2, 7, "=SUM(B3:G3)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet3, 3, 0, "East", null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 1, 45, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 2, 75, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 3, 50, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 4, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 5, 75, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 3, 6, 100, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet3, 3, 7, "=SUM(B4:G4)", null);

    _ = xlsxwriter.worksheet_write_string(worksheet3, 4, 0, "West", null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 1, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 2, 15, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 3, 55, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 4, 35, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 5, 20, null);
    _ = xlsxwriter.worksheet_write_number(worksheet3, 4, 6, 50, null);
    _ = xlsxwriter.worksheet_write_formula(worksheet3, 4, 7, "=SUM(B5:G5)", null);

    _ = xlsxwriter.worksheet_write_formula(worksheet3, 5, 7, "=SUM(H2:H5)", bold);

    // Add bold format to the first row.
    _ = xlsxwriter.worksheet_set_row(worksheet3, 0, xlsxwriter.LXW_DEF_ROW_HEIGHT, bold);

    // Set column formatting and the outline level.
    _ = xlsxwriter.worksheet_set_column(worksheet3, 0, 0, 10, bold);
    _ = xlsxwriter.worksheet_set_column_opt(worksheet3, 1, 6, 5, null, &options6);
    _ = xlsxwriter.worksheet_set_column(worksheet3, 7, 7, 10, null);

    // Example 4: Show all possible outline levels.
    var level1 = xlsxwriter.lxw_row_col_options{ .level = 1, .hidden = 0, .collapsed = 0 };
    var level2 = xlsxwriter.lxw_row_col_options{ .level = 2, .hidden = 0, .collapsed = 0 };
    var level3 = xlsxwriter.lxw_row_col_options{ .level = 3, .hidden = 0, .collapsed = 0 };
    var level4 = xlsxwriter.lxw_row_col_options{ .level = 4, .hidden = 0, .collapsed = 0 };
    var level5 = xlsxwriter.lxw_row_col_options{ .level = 5, .hidden = 0, .collapsed = 0 };
    var level6 = xlsxwriter.lxw_row_col_options{ .level = 6, .hidden = 0, .collapsed = 0 };
    var level7 = xlsxwriter.lxw_row_col_options{ .level = 7, .hidden = 0, .collapsed = 0 };

    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 0, "Level 1", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 1, 0, "Level 2", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 2, 0, "Level 3", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 3, 0, "Level 4", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 4, 0, "Level 5", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 5, 0, "Level 6", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 6, 0, "Level 7", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 7, 0, "Level 6", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 8, 0, "Level 5", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 9, 0, "Level 4", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 10, 0, "Level 3", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 11, 0, "Level 2", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 12, 0, "Level 1", null);

    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        0,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level1,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        1,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level2,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        2,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        3,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level4,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        4,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level5,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        5,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level6,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        6,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level7,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        7,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level6,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        8,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level5,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        9,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level4,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        10,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level3,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        11,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level2,
    );
    _ = xlsxwriter.worksheet_set_row_opt(
        worksheet4,
        12,
        xlsxwriter.LXW_DEF_ROW_HEIGHT,
        null,
        &level1,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
