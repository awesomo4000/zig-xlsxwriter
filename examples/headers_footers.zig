//
// This program shows several examples of how to set up headers and
// footers with libxlsxwriter.
//
// The control characters used in the header/footer strings are:
//
// Control             Category            Description
// =======             ========            ===========
// &L                  Justification       Left
// &C                                      Center
// &R                                      Right
//
// &P                  Information         Page number
// &N                                      Total number of pages
// &D                                      Date
// &T                                      Time
// &F                                      File name
// &A                                      Worksheet name
//
// &fontsize           Font                Font size
// &"font,style"                           Font name and style
// &U                                      Single underline
// &E                                      Double underline
// &S                                      Strikethrough
// &X                                      Superscript
// &Y                                      Subscript
//
// &[Picture]          Images              Image placeholder
// &G                                      Same as &[Picture]
//
// &&                  Miscellaneous       Literal ampersand &
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const mktmp = @import("mktmp");

// Embed the logo image directly into the executable
const logo_data = @embedFile("logo_small.png");

pub fn main() !void {
    // Create a temporary file for the logo using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(
        std.heap.page_allocator,
    );
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try mktmp.TmpFile.create(
        allocator,
        "logo_small_",
    );
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(logo_data);

    const workbook = xlsxwriter.workbook_new("zig-headers_footers.xlsx");

    const preview = "Select Print Preview to see the header and footer";

    // A simple example to start
    const worksheet1 = xlsxwriter.workbook_add_worksheet(workbook, "Simple");
    const header1 = "&CHere is some centered text.";
    const footer1 = "&LHere is some left aligned text.";

    _ = xlsxwriter.worksheet_set_header(worksheet1, header1);
    _ = xlsxwriter.worksheet_set_footer(worksheet1, footer1);

    _ = xlsxwriter.worksheet_set_column(worksheet1, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet1, 0, 0, preview, null);

    // An example with an image
    const worksheet2 = xlsxwriter.workbook_add_worksheet(workbook, "Image");
    var header_options = xlsxwriter.lxw_header_footer_options{
        .image_left = @as([*c]const u8, @ptrCast(tmp_file.path.ptr)),
        .image_center = null,
        .image_right = null,
    };

    _ = xlsxwriter.worksheet_set_header_opt(worksheet2, "&L&[Picture]", &header_options);

    _ = xlsxwriter.worksheet_set_margins(worksheet2, -1, -1, 1.3, -1);
    _ = xlsxwriter.worksheet_set_column(worksheet2, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 0, preview, null);

    // This is an example of some of the header/footer variables
    const worksheet3 = xlsxwriter.workbook_add_worksheet(workbook, "Variables");
    const header3 = "&LPage &P of &N" ++ "&CFilename: &F" ++ "&RSheetname: &A";
    const footer3 = "&LCurrent date: &D" ++ "&RCurrent time: &T";
    var breaks = [_]xlsxwriter.lxw_row_t{ 20, 0 };

    _ = xlsxwriter.worksheet_set_header(worksheet3, header3);
    _ = xlsxwriter.worksheet_set_footer(worksheet3, footer3);

    _ = xlsxwriter.worksheet_set_column(worksheet3, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 0, preview, null);

    _ = xlsxwriter.worksheet_set_h_pagebreaks(worksheet3, @ptrCast(&breaks));
    _ = xlsxwriter.worksheet_write_string(worksheet3, 20, 0, "Next page", null);

    // This example shows how to use more than one font
    const worksheet4 = xlsxwriter.workbook_add_worksheet(workbook, "Mixed fonts");
    const header4 = "&C&\"Courier New,Bold\"Hello &\"Arial,Italic\"World";
    const footer4 = "&C&\"Symbol\"e&\"Arial\" = mc&X2";

    _ = xlsxwriter.worksheet_set_header(worksheet4, header4);
    _ = xlsxwriter.worksheet_set_footer(worksheet4, footer4);

    _ = xlsxwriter.worksheet_set_column(worksheet4, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 0, preview, null);

    // Example of line wrapping
    const worksheet5 = xlsxwriter.workbook_add_worksheet(workbook, "Word wrap");
    const header5 = "&CHeading 1\nHeading 2";

    _ = xlsxwriter.worksheet_set_header(worksheet5, header5);

    _ = xlsxwriter.worksheet_set_column(worksheet5, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 0, 0, preview, null);

    // Example of inserting a literal ampersand &
    const worksheet6 = xlsxwriter.workbook_add_worksheet(workbook, "Ampersand");
    const header6 = "&CCuriouser && Curiouser - Attorneys at Law";

    _ = xlsxwriter.worksheet_set_header(worksheet6, header6);

    _ = xlsxwriter.worksheet_set_column(worksheet6, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 0, 0, preview, null);

    _ = xlsxwriter.workbook_close(workbook);
}
