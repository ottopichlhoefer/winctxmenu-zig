const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{
        .default_target = .{
            .cpu_arch = .x86_64,
            .os_tag = .windows,
        },
    });
    const optimize = b.standardOptimizeOption(.{});

    const exe = b.addExecutable(.{
        .name = "winctxmenu",
        .root_module = b.createModule(.{
            .root_source_file = b.path("src/main.zig"),
            .target = target,
            .optimize = optimize,
        }),
    });

    exe.subsystem = .Windows;
    exe.root_module.linkSystemLibrary("ole32", .{});
    exe.root_module.linkSystemLibrary("shell32", .{});
    exe.root_module.linkSystemLibrary("user32", .{});
    exe.root_module.linkSystemLibrary("advapi32", .{});
    exe.root_module.linkSystemLibrary("uuid", .{});

    b.installArtifact(exe);

    const run_cmd = b.addRunArtifact(exe);
    if (b.args) |args| run_cmd.addArgs(args);

    const run_step = b.step("run", "Run winctxmenu");
    run_step.dependOn(&run_cmd.step);
}
