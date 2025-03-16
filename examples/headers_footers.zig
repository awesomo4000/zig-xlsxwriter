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

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-headers_footers.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // TODO: Add specific example code here
    // Refer to the original C code for implementation details

    _ = xlsxwriter.workbook_close(workbook);
}
