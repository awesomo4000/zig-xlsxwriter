//
// An example of merging cells containing a rich string using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-merge_rich_string.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Configure a format for the merged range.
    const merge_format = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_align(merge_format, xlsxwriter.LXW_ALIGN_CENTER);
    _ = xlsxwriter.format_set_align(
        merge_format,
        xlsxwriter.LXW_ALIGN_VERTICAL_CENTER,
    );
    _ = xlsxwriter.format_set_border(merge_format, xlsxwriter.LXW_BORDER_THIN);

    // Configure formats for the rich string.
    const red = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(red, xlsxwriter.LXW_COLOR_RED);

    const blue = xlsxwriter.workbook_add_format(workbook);
    _ = xlsxwriter.format_set_font_color(blue, xlsxwriter.LXW_COLOR_BLUE);

    // Create the fragments for the rich string.
    var fragment1 = xlsxwriter.lxw_rich_string_tuple{
        .format = null,
        .string = "This is ",
    };
    var fragment2 = xlsxwriter.lxw_rich_string_tuple{
        .format = red,
        .string = "red",
    };
    var fragment3 = xlsxwriter.lxw_rich_string_tuple{
        .format = null,
        .string = " and this is ",
    };
    var fragment4 = xlsxwriter.lxw_rich_string_tuple{
        .format = blue,
        .string = "blue",
    };

    var rich_string = [_:null]?*xlsxwriter.lxw_rich_string_tuple{
        &fragment1,
        &fragment2,
        &fragment3,
        &fragment4,
        null,
    };

    // Write an empty string to the merged range.
    _ = xlsxwriter.worksheet_merge_range(
        worksheet,
        1,
        1,
        4,
        3,
        "",
        merge_format,
    );

    // We then overwrite the first merged cell with a rich string. Note that
    // we must also pass the cell format used in the merged cells format at
    // the end.
    _ = xlsxwriter.worksheet_write_rich_string(
        worksheet,
        1,
        1,
        &rich_string,
        merge_format,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
