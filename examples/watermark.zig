// An example of adding a worksheet watermark image using libxlsxwriter. This
// is based on the method of putting an image in the worksheet header as
// suggested in the Microsoft documentation:
// https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const mktmp = @import("mktmp");

const watermarkImageData = @embedFile("watermark.png");

pub fn main() !void {
    var arena = std.heap.ArenaAllocator.init(
        std.heap.page_allocator,
    );
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try mktmp.TmpFile.create(
        allocator,
        "watermark_",
    );

    defer tmp_file.cleanUp();

    try tmp_file.write(watermarkImageData);

    // Set the background using the temporary file
    // Convert the path to a null-terminated C string pointer

    const c_path = @as(
        [*c]const u8,
        @ptrCast(tmp_file.path.ptr),
    );
    const workbook =
        xlsxwriter.workbook_new(
            "zig-watermark.xlsx",
        );
    const worksheet =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );

    // Set a worksheet header with the watermark image.
    var header_options =
        xlsxwriter.lxw_header_footer_options{
            .image_left = null,
            .image_center = c_path,
            .image_right = null,
        };

    _ = xlsxwriter.worksheet_set_header_opt(
        worksheet,
        "&C&[Picture]",
        &header_options,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
