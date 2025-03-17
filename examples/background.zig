//
// An example of setting a worksheet background image with libxlsxwriter.
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

    // Create the workbook and add a worksheet
    const workbook =
        xlsxwriter.workbook_new(
            "zig-background.xlsx",
        );

    const worksheet =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );

    // Set the background using the temporary file
    // Convert the path to a null-terminated C string pointer
    const c_path = @as(
        [*c]const u8,
        @ptrCast(tmp_file.path.ptr),
    );
    _ = xlsxwriter.worksheet_set_background(
        worksheet,
        c_path,
    );

    // Close the workbook
    _ = xlsxwriter.workbook_close(workbook);

    // The temporary file will be automatically cleaned up by the defer statement
}
