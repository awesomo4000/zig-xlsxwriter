// /*
//  * Example of using libxlsxwriter to write a workbook file to a memory buffer.
//  *
//  * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//  *
//  */

// #include <stdio.h>

// #include "xlsxwriter.h"

// int main() {
//     const char *output_buffer;
//     size_t output_buffer_size;

//     /* Set the worksheet options. */
//     lxw_workbook_options options = {.output_buffer = &output_buffer,
//                                     .output_buffer_size = &output_buffer_size,
//                                     .constant_memory = LXW_FALSE,
//                                     .tmpdir = NULL,
//                                     .use_zip64 = LXW_FALSE};

//     /* Create a new workbook with options. */
//     lxw_workbook  *workbook  = workbook_new_opt(NULL, &options);
//     lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

//     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
//     worksheet_write_number(worksheet, 1, 0, 123,     NULL);

//     lxw_error error = workbook_close(workbook);

//     if (error)
//         return error;

//     /* Do something with the XLSX data in the output buffer. */
//     FILE *file = fopen("output_buffer.xlsx", "wb");
//     fwrite(output_buffer, output_buffer_size, 1, file);
//     fclose(file);
//     free((void *)output_buffer);

//     return ferror(stdout);
// }

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    var output_buffer: [*c]const u8 = undefined;
    var output_buffer_size: usize = undefined;

    // Set the worksheet options.
    var options = xlsxwriter.lxw_workbook_options{
        .output_buffer = &output_buffer,
        .output_buffer_size = &output_buffer_size,
        .constant_memory = xlsxwriter.LXW_FALSE,
        .tmpdir = null,
        .use_zip64 = xlsxwriter.LXW_FALSE,
    };

    // Create a new workbook with options.
    const workbook = xlsxwriter.workbook_new_opt(null, &options);
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    _ = xlsxwriter.worksheet_write_string(worksheet, 0, 0, "Hello", null);
    _ = xlsxwriter.worksheet_write_number(worksheet, 1, 0, 123, null);

    const error_code = xlsxwriter.workbook_close(workbook);

    if (error_code != xlsxwriter.LXW_NO_ERROR) {
        return error.WorkbookCloseFailed;
    }

    // Do something with the XLSX data in the output buffer.
    const file = try std.fs.cwd().createFile("zig-output_buffer.xlsx", .{});
    defer file.close();

    try file.writeAll(output_buffer[0..output_buffer_size]);

    // Free the buffer
    std.c.free(@constCast(output_buffer));
}
