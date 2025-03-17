//
// An example of embedding images into a worksheet using the libxlsxwriter
// library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const mktmp = @import("mktmp");

// Embed the logo image directly into the executable
const logo_data = @embedFile("logo.png");

pub fn main() !void {
    // Create a temporary file for the logo using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(
        std.heap.page_allocator,
    );
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try mktmp.TmpFile.create(
        allocator,
        "logo_",
    );
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(logo_data);

    // Create a new workbook and add a worksheet
    const workbook = xlsxwriter.workbook_new("zig-embed_images.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Change some of the column widths for clarity
    // COLS("A:B") expands to first_col, last_col
    const first_col = xlsxwriter.lxw_name_to_col("A:B");
    const last_col = xlsxwriter.lxw_name_to_col_2("A:B");
    _ = xlsxwriter.worksheet_set_column(worksheet, first_col, last_col, 30, null);

    // Embed an image
    const row_a2 = xlsxwriter.lxw_name_to_row("A2");
    const col_a2 = xlsxwriter.lxw_name_to_col("A2");
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        row_a2,
        col_a2,
        "Embed an image in a cell:",
        null,
    );

    const row_b2 = xlsxwriter.lxw_name_to_row("B2");
    const col_b2 = xlsxwriter.lxw_name_to_col("B2");
    _ = xlsxwriter.worksheet_embed_image(
        worksheet,
        row_b2,
        col_b2,
        tmp_file.path.ptr,
    );

    // Make a row bigger and embed the image
    _ = xlsxwriter.worksheet_set_row(worksheet, 3, 72, null);

    const row_a4 = xlsxwriter.lxw_name_to_row("A4");
    const col_a4 = xlsxwriter.lxw_name_to_col("A4");
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        row_a4,
        col_a4,
        "Embed an image in a cell:",
        null,
    );

    const row_b4 = xlsxwriter.lxw_name_to_row("B4");
    const col_b4 = xlsxwriter.lxw_name_to_col("B4");
    _ = xlsxwriter.worksheet_embed_image(
        worksheet,
        row_b4,
        col_b4,
        tmp_file.path.ptr,
    );

    // Make a row bigger and embed the image
    _ = xlsxwriter.worksheet_set_row(worksheet, 5, 150, null);

    const row_a6 = xlsxwriter.lxw_name_to_row("A6");
    const col_a6 = xlsxwriter.lxw_name_to_col("A6");
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        row_a6,
        col_a6,
        "Embed an image in a cell:",
        null,
    );

    const row_b6 = xlsxwriter.lxw_name_to_row("B6");
    const col_b6 = xlsxwriter.lxw_name_to_col("B6");
    _ = xlsxwriter.worksheet_embed_image(
        worksheet,
        row_b6,
        col_b6,
        tmp_file.path.ptr,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
