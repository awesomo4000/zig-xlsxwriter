//
// Example of setting document properties such as Author, Title, etc., for an
// Excel spreadsheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook = xlsxwriter.workbook_new("zig-doc_properties.xlsx");
    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);

    // Create a properties structure and set some of the fields.
    var properties = xlsxwriter.lxw_doc_properties{
        .title = "This is an example spreadsheet",
        .subject = "With document properties",
        .author = "John McNamara",
        .manager = "Dr. Heinz Doofenshmirtz",
        .company = "of Wolves",
        .category = "Example spreadsheets",
        .keywords = "Sample, Example, Properties",
        .comments = "Created with libxlsxwriter",
        .status = "Quo",
    };

    // Set the properties in the workbook.
    _ = xlsxwriter.workbook_set_properties(workbook, &properties);

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
