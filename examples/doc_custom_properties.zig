//
// Example of setting custom document properties for an Excel spreadsheet
// using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-doc_custom_properties.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);
    var datetime = xlsxwriter.lxw_datetime{
        .year = 2016,
        .month = 12,
        .day = 12,
        .hour = 0,
        .min = 0,
        .sec = 0.0,
    };

    // Set some custom document properties in the workbook.
    _ = xlsxwriter.workbook_set_custom_property_string(
        workbook,
        "Checked by",
        "Eve",
    );
    _ = xlsxwriter.workbook_set_custom_property_datetime(
        workbook,
        "Date completed",
        &datetime,
    );
    _ = xlsxwriter.workbook_set_custom_property_number(
        workbook,
        "Document number",
        12345,
    );
    _ = xlsxwriter.workbook_set_custom_property_number(
        workbook,
        "Reference number",
        1.2345,
    );
    _ = xlsxwriter.workbook_set_custom_property_boolean(
        workbook,
        "Has Review",
        1,
    );
    _ = xlsxwriter.workbook_set_custom_property_boolean(
        workbook,
        "Signed off",
        0,
    );

    // Add some text to the file.
    _ = xlsxwriter.worksheet_set_column(worksheet, 0, 0, 50, null);
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        0,
        0,
        "Select 'Workbook Properties' to see properties.",
        null,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
