const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook =
        xlsxwriter.workbook_new(
            "zig-rich_strings.xlsx",
        );
    const worksheet =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );

    // Set up some formats to use.
    const bold =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_bold(bold);

    const italic =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_italic(italic);

    const red =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(
        red,
        xlsxwriter.LXW_COLOR_RED,
    );

    const blue =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(
        blue,
        xlsxwriter.LXW_COLOR_BLUE,
    );

    const center =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_align(
        center,
        xlsxwriter.LXW_ALIGN_CENTER,
    );

    const superscript =
        xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_script(
        superscript,
        xlsxwriter.LXW_FONT_SUPERSCRIPT,
    );

    // Make the first column wider for clarity.
    _ = xlsxwriter.worksheet_set_column(
        worksheet,
        0,
        0,
        30,
        null,
    );

    // Example 1: Bold and italic text
    // Write individual cells with appropriate formatting
    var fragment11 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = null,
            .string = "This is ",
        };
    var fragment12 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = bold,
            .string = "bold",
        };
    var fragment13 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = null,
            .string = " and this is ",
        };
    var fragment14 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = italic,
            .string = "italic",
        };
    var rich_string1 =
        [_:null]?*xlsxwriter.lxw_rich_string_tuple{
            &fragment11,
            &fragment12,
            &fragment13,
            &fragment14,
            null,
        };
    _ = xlsxwriter.worksheet_write_rich_string(
        worksheet,
        0,
        0,
        &rich_string1,
        null,
    );

    // Example 2: Red and blue text

    var fragment21 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = null,
            .string = "This is ",
        };
    var fragment22 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = red,
            .string = "red",
        };
    var fragment23 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = null,
            .string = " and this is ",
        };
    var fragment24 =
        xlsxwriter.lxw_rich_string_tuple{
            .format = blue,
            .string = "blue",
        };
    var rich_string2 =
        [_:null]?*xlsxwriter.lxw_rich_string_tuple{
            &fragment21,
            &fragment22,
            &fragment23,
            &fragment24,
            null,
        };
    _ = xlsxwriter.worksheet_write_rich_string(
        worksheet,
        2,
        0,
        &rich_string2,
        null,
    );

    // Example 3. A rich string plus cell formatting.
    var fragment31 = xlsxwriter.lxw_rich_string_tuple{
        .format = null,
        .string = "Some ",
    };
    var fragment32 = xlsxwriter.lxw_rich_string_tuple{
        .format = bold,
        .string = "bold text",
    };
    var fragment33 = xlsxwriter.lxw_rich_string_tuple{
        .format = null,
        .string = " centered",
    };
    var rich_string3 =
        [_:null]?*xlsxwriter.lxw_rich_string_tuple{
            &fragment31,
            &fragment32,
            &fragment33,
            null,
        };
    _ = xlsxwriter.worksheet_write_rich_string(
        worksheet,
        4,
        0,
        &rich_string3,
        center,
    );

    // Example 4: Math example with superscript
    var fragment41 = xlsxwriter.lxw_rich_string_tuple{
        .format = italic,
        .string = "j =k",
    };
    var fragment42 = xlsxwriter.lxw_rich_string_tuple{
        .format = superscript,
        .string = "(n-1)",
    };
    var rich_string4 =
        [_:null]?*xlsxwriter.lxw_rich_string_tuple{
            &fragment41,
            &fragment42,
            null,
        };

    _ = xlsxwriter.worksheet_write_rich_string(
        worksheet,
        6,
        0,
        &rich_string4,
        center,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
