//
// An example of writing cell comments to a worksheet using libxlsxwriter.
//
// Each of the worksheets demonstrates different features of cell comments.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-comments2.xlsx");
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet5 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet6 = xlsxwriter.workbook_add_worksheet(workbook, null);
    const worksheet7 = xlsxwriter.workbook_add_worksheet(workbook, null);

    const text_wrap = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_text_wrap(text_wrap);
    _ = xlsxwriter.format_set_align(text_wrap, xlsxwriter.LXW_ALIGN_VERTICAL_TOP);

    // Example 1: Simple cell comment without formatting
    _ = xlsxwriter.worksheet_set_column(worksheet1, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet1, 2, 50, null);

    _ = xlsxwriter.worksheet_write_string(
        worksheet1,
        2,
        2,
        "Hold the mouse over this cell to see the comment.",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet1, 2, 2, "This is a comment.");

    // Example 2: Visible and hidden comments
    _ = xlsxwriter.worksheet_set_column(worksheet2, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet2, 2, 50, null);

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        2,
        2,
        "This cell comment is visible.",
        text_wrap,
    );

    var options2 = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_VISIBLE,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet2, 2, 2, "Hello.", &options2);

    _ = xlsxwriter.worksheet_write_string(
        worksheet2,
        5,
        2,
        "This cell comment isn't visible until you pass the mouse over it (the default).",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet2, 5, 2, "Hello.");

    // Example 3: Worksheet level comment visibility
    _ = xlsxwriter.worksheet_set_column(worksheet3, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet3, 2, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet3, 5, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet3, 8, 50, null);

    _ = xlsxwriter.worksheet_show_comments(worksheet3);

    _ = xlsxwriter.worksheet_write_string(
        worksheet3,
        2,
        2,
        "This cell comment is visible, explicitly.",
        text_wrap,
    );

    var options3a = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_VISIBLE,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet3, 2, 2, "Hello", &options3a);

    _ = xlsxwriter.worksheet_write_string(
        worksheet3,
        5,
        2,
        "This cell comment is also visible because we used worksheet_show_comments().",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet3, 5, 2, "Hello");

    _ = xlsxwriter.worksheet_write_string(
        worksheet3,
        8,
        2,
        "However, we can still override it locally.",
        text_wrap,
    );

    var options3b = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet3, 8, 2, "Hello", &options3b);

    // Example 4: Comment box dimensions
    _ = xlsxwriter.worksheet_set_column(worksheet4, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet4, 2, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet4, 5, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet4, 8, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet4, 15, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet4, 18, 50, null);

    _ = xlsxwriter.worksheet_show_comments(worksheet4);

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        2,
        2,
        "This cell comment is default size.",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet4, 2, 2, "Hello");

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        5,
        2,
        "This cell comment is twice as wide.",
        text_wrap,
    );

    var options4a = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 2.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet4, 5, 2, "Hello", &options4a);

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        8,
        2,
        "This cell comment is twice as high.",
        text_wrap,
    );

    var options4b = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 2.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet4, 8, 2, "Hello", &options4b);

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        15,
        2,
        "This cell comment is scaled in both directions.",
        text_wrap,
    );

    var options4c = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.2,
        .y_scale = 0.5,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet4, 15, 2, "Hello", &options4c);

    _ = xlsxwriter.worksheet_write_string(
        worksheet4,
        18,
        2,
        "This cell comment has width and height specified in pixels.",
        text_wrap,
    );

    var options4d = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 200,
        .height = 50,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet4, 18, 2, "Hello", &options4d);

    // Example 5: Comment positioning
    _ = xlsxwriter.worksheet_set_column(worksheet5, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet5, 2, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet5, 5, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet5, 8, 50, null);

    _ = xlsxwriter.worksheet_show_comments(worksheet5);

    _ = xlsxwriter.worksheet_write_string(
        worksheet5,
        2,
        2,
        "This cell comment is in the default position.",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet5, 2, 2, "Hello");

    _ = xlsxwriter.worksheet_write_string(
        worksheet5,
        5,
        2,
        "This cell comment has been moved to another cell.",
        text_wrap,
    );

    var options5a = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 3,
        .start_col = 4,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet5, 5, 2, "Hello", &options5a);

    _ = xlsxwriter.worksheet_write_string(
        worksheet5,
        8,
        2,
        "This cell comment has been shifted within its default cell.",
        text_wrap,
    );

    var options5b = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 30,
        .y_offset = 12,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet5, 8, 2, "Hello", &options5b);

    // Example 6: Comment colors
    _ = xlsxwriter.worksheet_set_column(worksheet6, 2, 2, 25, null);
    _ = xlsxwriter.worksheet_set_row(worksheet6, 2, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet6, 5, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet6, 8, 50, null);

    _ = xlsxwriter.worksheet_show_comments(worksheet6);

    _ = xlsxwriter.worksheet_write_string(
        worksheet6,
        2,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    var options6a = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x008000, // Green
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet6, 2, 2, "Hello", &options6a);

    _ = xlsxwriter.worksheet_write_string(
        worksheet6,
        5,
        2,
        "This cell comment has the default color.",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet6, 5, 2, "Hello");

    _ = xlsxwriter.worksheet_write_string(
        worksheet6,
        8,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    var options6b = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0xFF6600,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet6, 8, 2, "Hello", &options6b);

    // Example 7: Comment author
    _ = xlsxwriter.worksheet_set_column(worksheet7, 2, 2, 30, null);
    _ = xlsxwriter.worksheet_set_row(worksheet7, 2, 50, null);
    _ = xlsxwriter.worksheet_set_row(worksheet7, 5, 60, null);

    _ = xlsxwriter.worksheet_write_string(
        worksheet7,
        2,
        2,
        "Move the mouse over this cell and you will see 'Cell C3 " ++
            "commented by' (blank) in the status bar at the bottom.",
        text_wrap,
    );
    _ = xlsxwriter.worksheet_write_comment(worksheet7, 2, 2, "Hello");

    _ = xlsxwriter.worksheet_write_string(
        worksheet7,
        5,
        2,
        "Move the mouse over this cell and you will see 'Cell C6 " ++
            "commented by libxlsxwriter' in the status bar at the bottom.",
        text_wrap,
    );

    var options7a = xlsxwriter.lxw_comment_options{
        .visible = xlsxwriter.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = "libxlsxwriter",
        .width = 0,
        .height = 0,
    };
    _ = xlsxwriter.worksheet_write_comment_opt(worksheet7, 5, 2, "Hello", &options7a);

    _ = xlsxwriter.workbook_close(workbook);
}
