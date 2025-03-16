const std = @import("std");
const fs = std.fs;
const os = std.os;
const builtin = @import("builtin");
const crypto = std.crypto;
const mem = std.mem;
const process = std.process;

/// A temporary file that automatically handles cleanup
pub const TmpFile = struct {
    file: fs.File,
    path: []const u8,
    allocator: mem.Allocator,

    /// Creates a new temporary file with the given prefix
    pub fn create(
        allocator: mem.Allocator,
        prefix: []const u8,
    ) !TmpFile {
        const result = try tmpFile(allocator, prefix);
        return TmpFile{
            .file = result.file,
            .path = result.path,
            .allocator = allocator,
        };
    }

    /// Writes data to the temporary file
    pub fn write(self: *TmpFile, data: []const u8) !void {
        try self.file.writeAll(data);
    }

    /// Reads the entire content of the temporary file
    pub fn readAll(
        self: *TmpFile,
        buffer: []u8,
    ) ![]u8 {
        try self.file.seekTo(0);
        const bytes_read = try self.file.readAll(buffer);
        return buffer[0..bytes_read];
    }

    /// Closes the file and deletes it, freeing all resources
    pub fn cleanUp(self: *TmpFile) void {
        self.file.close();
        fs.deleteFileAbsolute(self.path) catch {};
        self.allocator.free(self.path);
    }
};

/// Creates a unique temporary file with the given prefix.
/// The file will be created in:
/// - TMP environment variable path if defined and exists
/// - TEMP environment variable path if on Windows
/// - /tmp if on POSIX systems
///
/// Returns the full path to the created temporary file.
/// Caller is responsible for closing the file and deleting it when done.
pub fn tmpFile(
    allocator: mem.Allocator,
    prefix: []const u8,
) !struct { file: fs.File, path: []const u8 } {
    // Get the temporary directory path
    const tmp_dir_path = try getTmpDir(allocator);
    defer allocator.free(tmp_dir_path);

    // Generate a unique filename with the prefix and 8 random characters
    const unique_path = try generateUniquePath(
        allocator,
        tmp_dir_path,
        prefix,
    );
    errdefer allocator.free(unique_path);

    // Create and open the file
    const file = try fs.createFileAbsolute(
        unique_path,
        .{ .read = true },
    );
    errdefer file.close();

    return .{
        .file = file,
        .path = unique_path,
    };
}

/// Gets the appropriate temporary directory path based on environment and platform
fn getTmpDir(allocator: mem.Allocator) ![]const u8 {
    // Platform-specific behavior
    if (builtin.os.tag == .windows) {
        // On Windows, check TEMP environment variable first
        if (process.getEnvVarOwned(
            allocator,
            "TEMP",
        )) |temp_path| {
            if (dirExists(temp_path)) {
                return temp_path; // Already allocated by getEnvVarOwned
            }
            allocator.free(temp_path);
        } else |_| {}

        // Fall back to TMP if TEMP doesn't exist
        if (process.getEnvVarOwned(
            allocator,
            "TMP",
        )) |tmp_path| {
            if (dirExists(tmp_path)) {
                return tmp_path; // Already allocated by getEnvVarOwned
            }
            allocator.free(tmp_path);
        } else |_| {}

        // If neither exists, use current directory as last resort
        return allocator.dupe(u8, ".");
    } else {
        // For non-Windows platforms, check TMP first
        if (process.getEnvVarOwned(
            allocator,
            "TMP",
        )) |tmp_path| {
            if (dirExists(tmp_path)) {
                return tmp_path; // Already allocated by getEnvVarOwned
            }
            allocator.free(tmp_path);
        } else |_| {}

        // On POSIX systems, use /tmp
        return allocator.dupe(u8, "/tmp");
    }
}

/// Checks if a directory exists
fn dirExists(path: []const u8) bool {
    fs.cwd().access(path, .{}) catch return false;
    return true;
}

/// Generates a unique path by combining the directory, prefix, and random characters
fn generateUniquePath(
    allocator: mem.Allocator,
    dir_path: []const u8,
    prefix: []const u8,
) ![]const u8 {
    // Generate 8 random characters (using 4 random bytes converted to 8 hex chars)
    var random_bytes: [4]u8 = undefined;
    crypto.random.bytes(&random_bytes);

    // Convert to lowercase hex string
    var random_hex: [8]u8 = undefined;
    _ = try std.fmt.bufPrint(
        &random_hex,
        "{x:0>8}",
        .{std.mem.readInt(u32, &random_bytes, .big)},
    );

    // Ensure all characters are lowercase (though fmt with {x} already does this)
    for (&random_hex) |*c| {
        c.* = std.ascii.toLower(c.*);
    }

    // Combine directory, prefix, and random string
    const path_separator =
        if (builtin.os.tag == .windows) "\\" else "/";

    return std.fmt.allocPrint(
        allocator,
        "{s}{s}{s}{s}",
        .{ dir_path, path_separator, prefix, random_hex },
    );
}

test "basic tmpFile usage" {
    // Basic usage test
    const allocator = std.testing.allocator;

    const result = try tmpFile(allocator, "test_");
    defer {
        result.file.close();
        fs.deleteFileAbsolute(result.path) catch {};
        allocator.free(result.path);
    }

    // Write some data to the file
    try result.file.writeAll("Hello, temporary file!");

    // Verify the file exists and has the correct content
    try result.file.seekTo(0);
    var buffer: [100]u8 = undefined;
    const bytes_read = try result.file.readAll(&buffer);
    try std.testing.expectEqualStrings(
        "Hello, temporary file!",
        buffer[0..bytes_read],
    );
}

test "TmpFile interface" {
    const allocator = std.testing.allocator;

    // Create a temporary file using the interface
    var tmp = try TmpFile.create(allocator, "test_");
    defer tmp.cleanUp();

    // Write data to the file
    try tmp.write("Hello from the interface!");

    // Read the data back
    var buffer: [100]u8 = undefined;
    const content = try tmp.readAll(&buffer);

    // Verify the content
    try std.testing.expectEqualStrings(
        "Hello from the interface!",
        content,
    );
}
