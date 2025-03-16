// An example of adding a worksheet watermark image using libxlsxwriter. This
// is based on the method of putting an image in the worksheet header as
// suggested in the Microsoft documentation:
// https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-watermark.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Set a worksheet header with the watermark image.
    var header_options = xlsxwriter.lxw_header_footer_options{
        .image_left = null,
        .image_center = "watermark.png",
        .image_right = null,
    };

    _ = xlsxwriter.worksheet_set_header_opt(worksheet, "&C&[Picture]", &header_options);

    _ = xlsxwriter.workbook_close(workbook);
}
