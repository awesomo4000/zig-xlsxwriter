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
const c = @cImport({
    @cInclude("string.h");
});
pub fn main() !void {
    const workbook =
        xlsxwriter.workbook_new(
            "zig-tables.xlsx",
        );
    const worksheet1 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    try write_worksheet_1(worksheet1);

    const worksheet2 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet3 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet4 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet5 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet6 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet7 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet8 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet9 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet10 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet11 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet12 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );
    const worksheet13 =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );

    const currency_format =
        xlsxwriter.workbook_add_format(workbook);
    xlsxwriter.format_set_num_format(
        currency_format,
        "$#,##0",
    );

    // Example 2: Default table with data
    _ = xlsxwriter.worksheet_set_column(
        worksheet2,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        0,
        1,
        "Default table with data.",
        null,
    );
    try write_worksheet_data(worksheet2, null);
    _ = xlsxwriter.worksheet_add_table(
        worksheet2,
        2,
        1,
        6,
        5,
        null,
    );

    // Example 3: Table without autofilter
    _ = xlsxwriter.worksheet_set_column(
        worksheet3,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet3,
        0,
        1,
        "Table without autofilter.",
        null,
    );

    try write_worksheet_data(worksheet3, null);

    var options3 = xlsxwriter.lxw_table_options{
        .no_autofilter = 1,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = null,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet3,
        2,
        1,
        6,
        5,
        &options3,
    );

    // Example 4: Table without default header row
    _ = xlsxwriter.worksheet_set_column(
        worksheet4,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        0,
        1,
        "Table without default header row.",
        null,
    );
    try write_worksheet_data(worksheet4, null);
    var options4 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 1,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = null,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet4,
        3,
        1,
        6,
        5,
        &options4,
    );

    // Example 5: Default table with "First Column" and "Last Column" options
    _ = xlsxwriter.worksheet_set_column(
        worksheet5,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet5,
        0,
        1,
        "Default table with \"First Column\" and \"Last Column\" options.",
        null,
    );
    try write_worksheet_data(worksheet5, null);
    var options5 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 1,
        .last_column = 1,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = null,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet5,
        2,
        1,
        6,
        5,
        &options5,
    );

    // Example 6: Table with banded columns but without default banded rows
    _ = xlsxwriter.worksheet_set_column(
        worksheet6,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet6,
        0,
        1,
        "Table with banded columns but without default banded rows.",
        null,
    );
    try write_worksheet_data(worksheet6, null);
    var options6 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 1,
        .banded_columns = 1,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = null,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet6,
        2,
        1,
        6,
        5,
        &options6,
    );

    // Example 7: Table with user defined column headers
    _ = xlsxwriter.worksheet_set_column(
        worksheet7,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet7,
        0,
        1,
        "Table with user defined column headers.",
        null,
    );
    try write_worksheet_data(worksheet7, null);
    var col7_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col7_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col7_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col7_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col7_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var columns7 = [_]?*xlsxwriter.lxw_table_column{
        &col7_1,
        &col7_2,
        &col7_3,
        &col7_4,
        &col7_5,
        null,
    };
    var options7 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = &columns7,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet7,
        2,
        1,
        6,
        5,
        &options7,
    );

    // Example 8: Table with user defined column headers and formula
    _ = xlsxwriter.worksheet_set_column(
        worksheet8,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet8,
        0,
        1,
        "Table with user defined column headers.",
        null,
    );
    try write_worksheet_data(worksheet8, null);
    var col8_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col8_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col8_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col8_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col8_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col8_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table8[@[Quarter 1]:[Quarter 4]])",
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var columns8 = [_]?*xlsxwriter.lxw_table_column{
        &col8_1,
        &col8_2,
        &col8_3,
        &col8_4,
        &col8_5,
        &col8_6,
        null,
    };
    var options8 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 0,
        .columns = &columns8,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet8,
        2,
        1,
        6,
        6,
        &options8,
    );

    // Example 9: Table with totals row (but no caption or totals)
    _ = xlsxwriter.worksheet_set_column(
        worksheet9,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet9,
        0,
        1,
        "Table with totals row (but no caption or totals).",
        null,
    );
    try write_worksheet_data(worksheet9, null);
    var col9_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col9_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col9_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col9_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col9_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col9_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table9[@[Quarter 1]:[Quarter 4]])",
        .format = null,
        .total_string = null,
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var columns9 = [_]?*xlsxwriter.lxw_table_column{
        &col9_1,
        &col9_2,
        &col9_3,
        &col9_4,
        &col9_5,
        &col9_6,
        null,
    };
    var options9 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 1,
        .columns = &columns9,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet9,
        2,
        1,
        7,
        6,
        &options9,
    );

    // Example 10: Table with totals row with user captions and functions
    _ = xlsxwriter.worksheet_set_column(
        worksheet10,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet10,
        0,
        1,
        "Table with totals row with user captions and functions.",
        null,
    );
    try write_worksheet_data(worksheet10, null);
    var col10_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = "Totals",
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col10_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col10_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col10_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col10_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col10_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table10[@[Quarter 1]:[Quarter 4]])",
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var columns10 = [_]?*xlsxwriter.lxw_table_column{
        &col10_1,
        &col10_2,
        &col10_3,
        &col10_4,
        &col10_5,
        &col10_6,
        null,
    };
    var options10 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 1,
        .columns = &columns10,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet10,
        2,
        1,
        7,
        6,
        &options10,
    );

    // Example 11: Table with alternative Excel style
    _ = xlsxwriter.worksheet_set_column(
        worksheet11,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet11,
        0,
        1,
        "Table with alternative Excel style.",
        null,
    );
    try write_worksheet_data(worksheet11, null);
    var col11_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = "Totals",
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col11_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col11_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col11_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col11_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col11_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table11[@[Quarter 1]:[Quarter 4]])",
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var columns11 = [_]?*xlsxwriter.lxw_table_column{
        &col11_1,
        &col11_2,
        &col11_3,
        &col11_4,
        &col11_5,
        &col11_6,
        null,
    };
    var options11 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = xlsxwriter.LXW_TABLE_STYLE_TYPE_LIGHT,
        .style_type_number = 11,
        .total_row = 1,
        .columns = &columns11,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet11,
        2,
        1,
        7,
        6,
        &options11,
    );

    // Example 12: Table with Excel style removed
    _ = xlsxwriter.worksheet_set_column(
        worksheet12,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet12,
        0,
        1,
        "Table with Excel style removed.",
        null,
    );
    try write_worksheet_data(worksheet12, null);
    var col12_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = "Totals",
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col12_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col12_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col12_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col12_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col12_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table12[@[Quarter 1]:[Quarter 4]])",
        .format = null,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var columns12 = [_]?*xlsxwriter.lxw_table_column{
        &col12_1,
        &col12_2,
        &col12_3,
        &col12_4,
        &col12_5,
        &col12_6,
        null,
    };
    var options12 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = xlsxwriter.LXW_TABLE_STYLE_TYPE_LIGHT,
        .style_type_number = 0,
        .total_row = 1,
        .columns = &columns12,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet12,
        2,
        1,
        7,
        6,
        &options12,
    );

    // Example 13: Table with column formats
    _ = xlsxwriter.worksheet_set_column(
        worksheet13,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet13,
        0,
        1,
        "Table with column formats.",
        null,
    );
    try write_worksheet_data(worksheet13, currency_format);
    var col13_1 = xlsxwriter.lxw_table_column{
        .header = "Product",
        .formula = null,
        .format = null,
        .total_string = "Totals",
        .total_function = 0,
        .total_value = 0,
        .header_format = null,
    };
    var col13_2 = xlsxwriter.lxw_table_column{
        .header = "Quarter 1",
        .formula = null,
        .format = currency_format,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col13_3 = xlsxwriter.lxw_table_column{
        .header = "Quarter 2",
        .formula = null,
        .format = currency_format,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col13_4 = xlsxwriter.lxw_table_column{
        .header = "Quarter 3",
        .formula = null,
        .format = currency_format,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col13_5 = xlsxwriter.lxw_table_column{
        .header = "Quarter 4",
        .formula = null,
        .format = currency_format,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var col13_6 = xlsxwriter.lxw_table_column{
        .header = "Year",
        .formula = "=SUM(Table13[@[Quarter 1]:[Quarter 4]])",
        .format = currency_format,
        .total_string = null,
        .total_function = xlsxwriter.LXW_TABLE_FUNCTION_SUM,
        .total_value = 0,
        .header_format = null,
    };
    var columns13 = [_]?*xlsxwriter.lxw_table_column{
        &col13_1,
        &col13_2,
        &col13_3,
        &col13_4,
        &col13_5,
        &col13_6,
        null,
    };
    var options13 = xlsxwriter.lxw_table_options{
        .no_autofilter = 0,
        .no_header_row = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = 0,
        .style_type_number = 0,
        .total_row = 1,
        .columns = &columns13,
        .name = null,
    };
    _ = xlsxwriter.worksheet_add_table(
        worksheet13,
        2,
        1,
        7,
        6,
        &options13,
    );
    _ = xlsxwriter.workbook_close(workbook);
}

fn write_worksheet_data(
    worksheet: ?*xlsxwriter.lxw_worksheet,
    format: ?*xlsxwriter.lxw_format,
) !void {
    // array of strings "Apples", "Pears", "Bananas", "Oranges"
    const rowDescriptions: [4][:0]const u8 =
        .{ "Apples", "Pears", "Bananas", "Oranges" };

    // data is a 4x4 array of f64
    const data: [4][4]f64 = [4][4]f64{
        [4]f64{ 10000, 5000, 8000, 6000 },
        [4]f64{ 2000, 3000, 4000, 5000 },
        [4]f64{ 6000, 6000, 6500, 6000 },
        [4]f64{ 500, 300, 200, 700 },
    };
    const startRow: usize = 3;

    for (rowDescriptions, 0..) |str, i| {
        // Write the first row strings
        _ = xlsxwriter.worksheet_write_string(
            worksheet,
            @intCast(i + startRow),
            1,
            str,
            format,
        );
    }

    for (data, 0..) |row, i| {
        for (row, 0..) |value, j| {
            _ = xlsxwriter.worksheet_write_number(
                worksheet,
                @intCast(i + startRow),
                @intCast(j + 2),
                value,
                format,
            );
        }
    }
}

fn write_worksheet_1(
    worksheet1: ?*xlsxwriter.lxw_worksheet,
) !void {

    // Example 1: Default table with no data
    _ = xlsxwriter.worksheet_set_column(
        worksheet1,
        1,
        6,
        12,
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        0,
        1,
        "Default table with no data.",
        null,
    );

    _ = xlsxwriter.worksheet_add_table(
        worksheet1,
        2,
        1,
        6,
        5,
        null,
    );
}
