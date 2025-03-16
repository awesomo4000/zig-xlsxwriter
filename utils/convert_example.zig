const std = @import("std");
const fs = std.fs;
const mem = std.mem;
const Allocator = mem.Allocator;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    // Parse command line arguments
    const args = try std.process.argsAlloc(allocator);
    defer std.process.argsFree(allocator, args);

    if (args.len < 2) {
        const prog_name = extractBaseName(args[0]);
        std.debug.print("Usage: {s} <input_file>\n", .{prog_name});
        std.debug.print("\nRead a .c libxlswriter example, write zig version to stdout\n", .{});
        return;
    }

    const input_file_path = args[1];

    // Read the input file
    const c_content = fs.cwd().readFileAlloc(allocator, input_file_path, 1024 * 1024) catch {
        std.debug.print("Error: Input file '{s}' not found or could not be read\n", .{input_file_path});
        return;
    };
    defer allocator.free(c_content);

    // Extract the base name from the input file path
    const base_name = extractBaseName(input_file_path);

    // Convert C to Zig
    const zig_content = try convertCToZig(allocator, c_content, base_name);
    defer allocator.free(zig_content);

    // Output the converted content to stdout
    const stdout = std.io.getStdOut().writer();
    try stdout.writeAll(zig_content);
}

fn extractBaseName(file_path: []const u8) []const u8 {
    // Find the last path separator
    const last_slash = mem.lastIndexOfScalar(u8, file_path, '/') orelse return file_path;

    // Extract the filename part
    const filename = file_path[last_slash + 1 ..];

    // Find the last dot (extension separator)
    const last_dot = mem.lastIndexOfScalar(u8, filename, '.') orelse return filename;

    // Return the base name without extension
    return filename[0..last_dot];
}

fn convertCToZig(allocator: Allocator, c_content: []const u8, base_name: []const u8) ![]u8 {
    // Extract comments from the beginning
    var comment_end: usize = 0;
    if (mem.startsWith(u8, c_content, "/*")) {
        if (mem.indexOf(u8, c_content, "*/")) |end| {
            comment_end = end + 2;
        }
    }

    // Convert C-style comments to Zig-style
    var comments = std.ArrayList(u8).init(allocator);
    defer comments.deinit();

    if (comment_end > 0) {
        var lines = mem.splitSequence(u8, c_content[0..comment_end], "\n");
        while (lines.next()) |line| {
            const trimmed = mem.trim(u8, line, " \t\r\n/*");
            if (trimmed.len > 0) {
                try comments.writer().print("// {s}\n", .{trimmed});
            } else {
                try comments.writer().print("//\n", .{});
            }
        }
    }

    // Extract the main function content
    var main_start: usize = 0;
    var main_end: usize = c_content.len;

    if (mem.indexOf(u8, c_content, "int main(")) |start| {
        main_start = start;
        var brace_count: usize = 0;
        var in_main = false;

        for (c_content[start..], 0..) |c, i| {
            if (c == '{') {
                brace_count += 1;
                in_main = true;
            } else if (c == '}') {
                brace_count -= 1;
                if (in_main and brace_count == 0) {
                    main_end = start + i + 1;
                    break;
                }
            }
        }
    }

    // Create the Zig file content
    var result = std.ArrayList(u8).init(allocator);
    defer result.deinit();

    try result.appendSlice(comments.items);
    try result.appendSlice("\nconst std = @import(\"std\");\n");
    try result.appendSlice("const xlsxwriter = @import(\"xlsxwriter\");\n\n");
    try result.appendSlice("pub fn main() !void {\n");

    // Add the main content with output path modified
    try result.writer().print("    const workbook = xlsxwriter.workbook_new(\"zig-{s}.xlsx\");\n", .{base_name});
    try result.appendSlice("    const worksheet = xlsxwriter.workbook_add_worksheet(workbook, null);\n\n");
    try result.appendSlice("    // TODO: Add specific example code here\n");
    try result.appendSlice("    // Refer to the original C code for implementation details\n\n");
    try result.appendSlice("    _ = xlsxwriter.workbook_close(workbook);\n");
    try result.appendSlice("}\n");

    return result.toOwnedSlice();
}
