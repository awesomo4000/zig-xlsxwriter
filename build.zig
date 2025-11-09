const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    // Add a clean step to remove zig-out and .zig-cache
    const clean_step = b.step(
        "clean",
        "Remove zig-out and .zig-cache directories",
    );
    const clean_cmd = b.addSystemCommand(&.{
        "rm", "-rf", "zig-out", ".zig-cache",
    });
    clean_step.dependOn(&clean_cmd.step);

    // Add an option for specifying which example to run
    const example_option = b.option(
        []const u8,
        "example",
        "Specify which example to run",
    );

    const xlsxwriter_dep = b.dependency("xlsxwriter", .{
        .target = target,
        .optimize = optimize,
        .USE_SYSTEM_MINIZIP = true,
    });
    const xlsxwriter_module = b.addModule("xlsxwriter", .{
        .root_source_file = b.path("src/xlsxwriter.zig"),
    });

    // Add mktmp module
    const mktmp_module = b.addModule("mktmp", .{
        .root_source_file = b.path("src/mktmp.zig"),
    });

    // get libxlsxwriter
    xlsxwriter_module.linkLibrary(xlsxwriter_dep.artifact("xlsxwriter"));
    xlsxwriter_module.link_libc = true;

    // Create a run step that depends on the example option
    const run_step = b.step(
        "run",
        "Run the specified example (use -Dexample=name)",
    );

    // Create a default step to build all examples
    const all_step = b.getInstallStep();

    // Check if examples/ directory exists. This is necessary to avoid warnings
    // when this is used as a dependency.
    const examples_dir = "examples";
    var has_examples = false;
    if (std.fs.cwd().access(examples_dir, .{})) {
        has_examples = true;
    } else |err| {
        if (err != error.FileNotFound) {
            std.debug.print(
                "[zig-xlsxwriter:WARN] Error checking examples directory: {}\n",
                .{err},
            );
        }
    }

    if (has_examples) {
        var dir = std.fs.cwd().openDir(
            examples_dir,
            .{ .iterate = true },
        ) catch |err| {
            std.debug.print(
                "[zig-xlsxwriter:WARN] Error opening examples directory: {}\n",
                .{err},
            );
            return;
        };

        defer dir.close();

        var iter = dir.iterate();
        while (iter.next() catch |err| {
            std.debug.print(
                "[zig-xlsxwriter:WARN] Error iterating examples directory: {}\n",
                .{err},
            );
            return;
        }) |entry| {
            if (entry.kind != .file) continue;

            // Only process .zig files
            if (!std.mem.endsWith(u8, entry.name, ".zig")) continue;

            const example_path = b.fmt(
                "{s}/{s}",
                .{ examples_dir, entry.name },
            );

            const install_step = makeExample(b, .{
                .path = example_path,
                .module = xlsxwriter_module,
                .mktmp_module = mktmp_module,
                .target = target,
                .optimize = optimize,
                .run_step = run_step,
                .example_option = example_option,
            });

            // Make the default step depend on building all examples
            all_step.dependOn(install_step);
        }
    }
}

fn getExampleName(filename: []const u8) []const u8 {
    var split =
        std.mem.splitSequence(u8, filename, ".");
    return split.first();
}

fn makeExample(b: *std.Build, options: BuildInfo) *std.Build.Step {
    const example_name = options.filename();

    // Create a build step that only builds and installs
    const build_step = b.step(
        example_name,
        b.fmt("Build the {s} example", .{example_name}),
    );

    const exe = b.addExecutable(.{
        .name = example_name,
        .root_source_file = b.path(options.path),
        .target = options.target,
        .optimize = options.optimize,
    });

    exe.root_module.addImport("xlsxwriter", options.module);
    exe.root_module.addImport("mktmp", options.mktmp_module);

    // Install the executable in zig-out/bin
    const install_exe = b.addInstallArtifact(exe, .{});

    // Make the build step depend on the installation
    build_step.dependOn(&install_exe.step);

    // Add run command to the run step if this example matches the option
    if (options.example_option) |requested_example| {
        if (std.mem.eql(u8, requested_example, example_name)) {
            const run_cmd = b.addRunArtifact(exe);
            options.run_step.dependOn(&run_cmd.step);
        }
    }

    return &install_exe.step;
}

const BuildInfo = struct {
    target: std.Build.ResolvedTarget,
    optimize: std.builtin.OptimizeMode,
    module: *std.Build.Module,
    mktmp_module: *std.Build.Module,
    path: []const u8,
    run_step: *std.Build.Step,
    example_option: ?[]const u8,

    fn filename(self: BuildInfo) []const u8 {
        var split = std.mem.splitSequence(
            u8,
            std.fs.path.basename(self.path),
            ".",
        );
        return split.first();
    }
};
