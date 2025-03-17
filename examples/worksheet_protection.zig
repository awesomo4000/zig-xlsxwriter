//
// Example of cell locking and formula hiding in an Excel
// worksheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    const workbook =
        xlsxwriter.workbook_new(
            "zig-worksheet_protection.xlsx",
        );
    const worksheet =
        xlsxwriter.workbook_add_worksheet(
            workbook,
            null,
        );

    const unlocked =
        xlsxwriter.workbook_add_format(
            workbook,
        );
    _ = xlsxwriter.format_set_unlocked(unlocked);

    const hidden =
        xlsxwriter.workbook_add_format(
            workbook,
        );
    _ = xlsxwriter.format_set_hidden(hidden);

    // Widen the first column to make the text clearer
    _ = xlsxwriter.worksheet_set_column(
        worksheet,
        0,
        0,
        40,
        null,
    );

    // Turn worksheet protection on without a password
    _ = xlsxwriter.worksheet_protect(
        worksheet,
        null,
        null,
    );

    // Write a locked, unlocked and hidden cell
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        0,
        0,
        "B1 is locked. It cannot be edited.",
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        1,
        0,
        "B2 is unlocked. It can be edited.",
        null,
    );
    _ = xlsxwriter.worksheet_write_string(
        worksheet,
        2,
        0,
        "B3 is hidden. The formula isn't visible.",
        null,
    );

    _ = xlsxwriter.worksheet_write_formula(
        worksheet,
        0,
        1,
        "=1+2",
        null,
    ); // Locked by default
    _ = xlsxwriter.worksheet_write_formula(
        worksheet,
        1,
        1,
        "=1+2",
        unlocked,
    );
    _ = xlsxwriter.worksheet_write_formula(
        worksheet,
        2,
        1,
        "=1+2",
        hidden,
    );

    _ = xlsxwriter.workbook_close(workbook);
}
