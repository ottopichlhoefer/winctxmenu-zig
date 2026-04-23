const std = @import("std");

const c = @cImport({
    @cDefine("COBJMACROS", "1");
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    @cDefine("WIN32_LEAN_AND_MEAN", "1");
    @cDefine("_WIN32_WINNT", "0x0A00");
    @cInclude("windows.h");
    @cInclude("objbase.h");
    @cInclude("shellapi.h");
    @cInclude("shlobj.h");
    @cInclude("shobjidl.h");
});

const win = std.os.windows;

extern "advapi32" fn RegCreateKeyExW(
    hKey: win.HKEY,
    lpSubKey: [*:0]const u16,
    Reserved: win.DWORD,
    lpClass: ?[*:0]u16,
    dwOptions: win.DWORD,
    samDesired: win.REGSAM,
    lpSecurityAttributes: ?*anyopaque,
    phkResult: *win.HKEY,
    lpdwDisposition: ?*win.DWORD,
) callconv(.winapi) win.LSTATUS;

extern "advapi32" fn RegSetValueExW(
    hKey: win.HKEY,
    lpValueName: ?[*:0]const u16,
    Reserved: win.DWORD,
    dwType: win.DWORD,
    lpData: ?[*]const u8,
    cbData: win.DWORD,
) callconv(.winapi) win.LSTATUS;

extern "advapi32" fn RegDeleteTreeW(
    hKey: win.HKEY,
    lpSubKey: [*:0]const u16,
) callconv(.winapi) win.LSTATUS;

extern "advapi32" fn RegCloseKey(
    hKey: win.HKEY,
) callconv(.winapi) win.LSTATUS;

// Per-monitor DPI awareness (Windows 10 1703+). The docs declare the parameter
// as DPI_AWARENESS_CONTEXT — a DECLARE_HANDLE pseudo-pointer whose permitted
// values are the negative integer constants -1..-5. Typed as isize here so we
// can pass -4 (PER_MONITOR_AWARE_V2) directly.
extern "user32" fn SetProcessDpiAwarenessContext(
    value: isize,
) callconv(.winapi) win.BOOL;

const DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2: isize = -4;

const Action = enum {
    show,
    install,
    uninstall,
    copy_hdrop,
    help,
};

const WindowMode = enum {
    hidden,
    auto,
};

const ParsedArgs = struct {
    allocator: std.mem.Allocator,
    action: Action = .show,
    mode_explicit: bool = false,
    window_mode: WindowMode = .hidden,
    x: ?i32 = null,
    y: ?i32 = null,
    paths: std.ArrayList([]const u8),

    fn init(allocator: std.mem.Allocator) ParsedArgs {
        return .{
            .allocator = allocator,
            .paths = .empty,
        };
    }

    fn deinit(self: *ParsedArgs) void {
        self.paths.deinit(self.allocator);
    }
};

const ComInit = struct {
    co_uninit: bool = false,
    ole_uninit: bool = false,

    fn deinit(self: ComInit) void {
        if (self.ole_uninit) c.OleUninitialize();
        if (self.co_uninit) c.CoUninitialize();
    }
};

const OwnerWindow = struct {
    hwnd: c.HWND,
    owned: bool,

    fn deinit(self: *OwnerWindow) void {
        if (self.owned) {
            _ = c.DestroyWindow(self.hwnd);
        }
    }
};

const Selection = struct {
    parent_folder: *c.IShellFolder,
    full_pidls: []?*c.ITEMIDLIST,
    child_pidls: []?*const c.ITEMIDLIST,

    fn init(allocator: std.mem.Allocator, absolute_paths: []const []const u8) !Selection {
        if (absolute_paths.len == 0) {
            return error.NoPaths;
        }

        var full_pidls = try allocator.alloc(?*c.ITEMIDLIST, absolute_paths.len);
        errdefer allocator.free(full_pidls);
        var child_pidls = try allocator.alloc(?*const c.ITEMIDLIST, absolute_paths.len);
        errdefer allocator.free(child_pidls);

        @memset(full_pidls, null);
        @memset(child_pidls, null);

        errdefer {
            for (full_pidls) |pidl| {
                if (pidl) |p| c.CoTaskMemFree(@ptrCast(p));
            }
            allocator.free(full_pidls);
            allocator.free(child_pidls);
        }

        const first_parent = std.fs.path.dirname(absolute_paths[0]) orelse absolute_paths[0];
        for (absolute_paths, 0..) |path, idx| {
            const parent = std.fs.path.dirname(path) orelse path;
            if (!std.ascii.eqlIgnoreCase(parent, first_parent)) {
                return error.MultipleParentFoldersNotSupported;
            }
            full_pidls[idx] = try parseDisplayNameToPidl(allocator, path);
        }

        var ppv_parent: ?*anyopaque = null;
        var child0: ?*const c.ITEMIDLIST = null;
        const hr_parent = c.SHBindToParent(full_pidls[0].?, &c.IID_IShellFolder, &ppv_parent, &child0);
        if (!succeeded(hr_parent) or ppv_parent == null or child0 == null) {
            return error.BindParentFailed;
        }

        const parent_folder: *c.IShellFolder = @ptrCast(@alignCast(ppv_parent.?));
        errdefer releaseIUnknown(parent_folder);

        child_pidls[0] = @ptrCast(child0.?);

        for (1..absolute_paths.len) |idx| {
            var temp_parent: ?*anyopaque = null;
            var child: ?*const c.ITEMIDLIST = null;
            const hr = c.SHBindToParent(full_pidls[idx].?, &c.IID_IShellFolder, &temp_parent, &child);
            if (!succeeded(hr) or child == null) {
                if (temp_parent != null) releaseCom(temp_parent);
                return error.BindParentFailed;
            }

            child_pidls[idx] = @ptrCast(child.?);
            if (temp_parent != null) releaseCom(temp_parent);
        }

        return .{
            .parent_folder = parent_folder,
            .full_pidls = full_pidls,
            .child_pidls = child_pidls,
        };
    }

    fn deinit(self: *Selection, allocator: std.mem.Allocator) void {
        releaseIUnknown(self.parent_folder);
        for (self.full_pidls) |pidl| {
            if (pidl) |p| c.CoTaskMemFree(@ptrCast(p));
        }
        allocator.free(self.full_pidls);
        allocator.free(self.child_pidls);
    }

    fn showMenu(self: *Selection, allocator: std.mem.Allocator, owner_hwnd: c.HWND, point: c.POINT, selected_paths: []const []const u8) !void {
        var context_menu_ppv: ?*anyopaque = null;
        const apidl: [*c]const ?*const c.ITEMIDLIST = @ptrCast(self.child_pidls.ptr);
        const folder_vtbl = self.parent_folder.lpVtbl orelse return error.GetUIObjectOfFailed;
        const get_ui_object_of = folder_vtbl.*.GetUIObjectOf orelse return error.GetUIObjectOfFailed;
        const hr_get_ui = get_ui_object_of(
            self.parent_folder,
            owner_hwnd,
            @as(c.UINT, @intCast(self.child_pidls.len)),
            @ptrCast(@constCast(apidl)),
            &c.IID_IContextMenu,
            null,
            &context_menu_ppv,
        );
        if (!succeeded(hr_get_ui) or context_menu_ppv == null) {
            return error.GetUIObjectOfFailed;
        }

        const context_menu: *c.IContextMenu = @ptrCast(@alignCast(context_menu_ppv.?));
        defer releaseIUnknown(context_menu);

        const hmenu = c.CreatePopupMenu();
        if (hmenu == null) {
            return error.CreatePopupMenuFailed;
        }
        defer _ = c.DestroyMenu(hmenu);

        const first_cmd: c.UINT = 1;
        const menu_vtbl = context_menu.lpVtbl orelse return error.QueryContextMenuFailed;
        const query_context_menu = menu_vtbl.*.QueryContextMenu orelse return error.QueryContextMenuFailed;
        const hr_qcm = query_context_menu(
            context_menu,
            hmenu,
            0,
            first_cmd,
            0x7FFF,
            c.CMF_NORMAL,
        );
        if (!succeeded(hr_qcm)) {
            return error.QueryContextMenuFailed;
        }

        attachMenuMsgForwarders(context_menu);
        defer clearMenuMsgForwarders();
        g_menu_command_id = 0;

        _ = c.SetForegroundWindow(owner_hwnd);
        const selected = c.TrackPopupMenuEx(
            hmenu,
            c.TPM_RETURNCMD | c.TPM_RIGHTBUTTON,
            point.x,
            point.y,
            owner_hwnd,
            null,
        );
        _ = c.PostMessageW(owner_hwnd, c.WM_NULL, 0, 0);
        diagLog(allocator, "TrackPopupMenuEx selected={d} at x={d} y={d}", .{ selected, point.x, point.y });

        var selected_cmd: c.UINT = 0;
        if (selected == 0) {
            if (g_menu_command_id != 0) {
                selected_cmd = g_menu_command_id;
                diagLog(allocator, "using WM_COMMAND fallback cmd={d}", .{selected_cmd});
            } else {
                diagLog(allocator, "menu dismissed or command not returned (selected=0)", .{});
                return;
            }
        } else {
            if (selected < first_cmd) return error.InvalidCommandId;
            selected_cmd = @intCast(selected);
        }
        if (selected_cmd < first_cmd) return error.InvalidCommandId;

        const verb_offset: usize = @intCast(selected_cmd - first_cmd);
        const verb_ptr_a: c.LPCSTR = @ptrFromInt(verb_offset);
        diagLog(allocator, "menu selected cmd={d} verb_offset={d}", .{ selected_cmd, verb_offset });

        if (isLikelyCopyCommand(allocator, menu_vtbl, context_menu, hmenu, selected_cmd, verb_offset)) {
            diagLog(allocator, "copy detected via probe, using CF_HDROP fallback for {d} path(s)", .{selected_paths.len});
            setClipboardFileDropList(allocator, selected_paths) catch |err| {
                diagLog(allocator, "setClipboardFileDropList error={s}", .{@errorName(err)});
                return err;
            };
            diagLog(allocator, "setClipboardFileDropList success", .{});
            return;
        }

        var invoke: c.CMINVOKECOMMANDINFOEX = std.mem.zeroes(c.CMINVOKECOMMANDINFOEX);
        invoke.cbSize = @sizeOf(c.CMINVOKECOMMANDINFOEX);
        invoke.fMask = c.CMIC_MASK_NOASYNC;
        invoke.hwnd = owner_hwnd;
        invoke.lpVerb = verb_ptr_a;
        invoke.nShow = c.SW_SHOWNORMAL;

        const invoke_command = menu_vtbl.*.InvokeCommand orelse return error.InvokeCommandFailed;
        const invoke_base: [*c]c.CMINVOKECOMMANDINFO = @ptrCast(&invoke);
        const seq_before = c.GetClipboardSequenceNumber();
        const hr_invoke = invoke_command(context_menu, invoke_base);
        diagLog(allocator, "InvokeCommand hr={d} seq_before={d}", .{ hr_invoke, seq_before });
        if (!succeeded(hr_invoke)) {
            return error.InvokeCommandFailed;
        }

        // Clipboard-related verbs such as "Copy" may use delayed rendering.
        // Keep pumping briefly so deferred providers run, then flush so data
        // survives after this short-lived process exits.
        waitForClipboardUpdate(seq_before);
        const seq_after = c.GetClipboardSequenceNumber();
        diagLog(allocator, "post-invoke seq_after={d}", .{seq_after});
        _ = c.OleFlushClipboard();
        diagLog(allocator, "OleFlushClipboard after invoke complete", .{});
    }
};

var g_context_menu2: ?*c.IContextMenu2 = null;
var g_context_menu3: ?*c.IContextMenu3 = null;
var g_menu_command_id: c.UINT = 0;

const hidden_window_class: c.LPCWSTR = std.unicode.utf8ToUtf16LeStringLiteral("WinCtxMenuHiddenOwnerWindow");

pub fn main() void {
    run() catch |err| {
        std.log.err("{s}", .{@errorName(err)});
        std.process.exit(1);
    };
}

fn run() !void {
    // Declare per-monitor DPI awareness so x/y arguments are interpreted as
    // physical screen pixels (matching what callers like Emacs report) rather
    // than being auto-scaled by Windows from a virtualized 96-DPI space.
    _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const args = try std.process.argsAlloc(allocator);
    defer std.process.argsFree(allocator, args);
    diagLog(allocator, "run args_count={d}", .{args.len});

    var parsed = try parseArgs(allocator, args);
    defer parsed.deinit();
    diagLog(allocator, "parsed action={s} paths={d}", .{ @tagName(parsed.action), parsed.paths.items.len });

    if (parsed.action == .help) {
        printUsage();
        return;
    }

    switch (parsed.action) {
        .install => try installRegistryEntries(allocator),
        .uninstall => try uninstallRegistryEntries(allocator),
        .show => try runShow(allocator, &parsed),
        .copy_hdrop => try runCopyHdrop(allocator, &parsed),
        .help => unreachable,
    }
}

fn parseArgs(allocator: std.mem.Allocator, args: []const []const u8) !ParsedArgs {
    var parsed = ParsedArgs.init(allocator);
    errdefer parsed.deinit();

    var i: usize = 1;
    while (i < args.len) : (i += 1) {
        const arg = args[i];

        if (std.mem.eql(u8, arg, "--help") or std.mem.eql(u8, arg, "-h")) {
            parsed.action = .help;
            continue;
        }
        if (std.mem.eql(u8, arg, "--install")) {
            try setAction(&parsed, .install);
            continue;
        }
        if (std.mem.eql(u8, arg, "--uninstall")) {
            try setAction(&parsed, .uninstall);
            continue;
        }
        if (std.mem.eql(u8, arg, "--copy-hdrop")) {
            try setAction(&parsed, .copy_hdrop);
            continue;
        }
        if (std.mem.eql(u8, arg, "--x")) {
            i += 1;
            if (i >= args.len) return error.MissingValueForX;
            parsed.x = try std.fmt.parseInt(i32, args[i], 10);
            continue;
        }
        if (std.mem.eql(u8, arg, "--y")) {
            i += 1;
            if (i >= args.len) return error.MissingValueForY;
            parsed.y = try std.fmt.parseInt(i32, args[i], 10);
            continue;
        }
        if (std.mem.eql(u8, arg, "--window")) {
            i += 1;
            if (i >= args.len) return error.MissingValueForWindow;
            if (std.mem.eql(u8, args[i], "hidden")) {
                parsed.window_mode = .hidden;
            } else if (std.mem.eql(u8, args[i], "auto")) {
                parsed.window_mode = .auto;
            } else {
                return error.InvalidWindowMode;
            }
            continue;
        }
        if (std.mem.startsWith(u8, arg, "-")) return error.UnknownFlag;

        try parsed.paths.append(allocator, arg);
    }

    if ((parsed.action == .install or parsed.action == .uninstall) and parsed.paths.items.len > 0) {
        return error.ModeDoesNotAcceptPaths;
    }
    if ((parsed.action == .install or parsed.action == .uninstall) and (parsed.x != null or parsed.y != null or parsed.window_mode != .hidden)) {
        return error.ModeDoesNotAcceptMenuFlags;
    }
    if (parsed.action == .copy_hdrop and parsed.paths.items.len == 0) {
        return error.CopyHdropNeedsPaths;
    }

    return parsed;
}

fn setAction(parsed: *ParsedArgs, action: Action) !void {
    if (!parsed.mode_explicit) {
        parsed.action = action;
        parsed.mode_explicit = true;
        return;
    }
    if (parsed.action != action) return error.ConflictingModes;
}

fn runShow(allocator: std.mem.Allocator, parsed: *const ParsedArgs) !void {
    const com = try initializeCom();
    defer com.deinit();

    const absolute_paths = try normalizePaths(allocator, parsed.paths.items);
    defer freePathList(allocator, absolute_paths);
    diagLog(allocator, "runShow absolute_paths={d}", .{absolute_paths.len});

    var selection = try Selection.init(allocator, absolute_paths);
    defer selection.deinit(allocator);

    var owner = try acquireOwnerWindow(parsed.window_mode);
    defer owner.deinit();

    const point = try resolvePoint(parsed.x, parsed.y);
    try selection.showMenu(allocator, owner.hwnd, point, absolute_paths);
}

fn runCopyHdrop(allocator: std.mem.Allocator, parsed: *const ParsedArgs) !void {
    const com = try initializeCom();
    defer com.deinit();

    const absolute_paths = try normalizePaths(allocator, parsed.paths.items);
    defer freePathList(allocator, absolute_paths);

    try setClipboardFileDropList(allocator, absolute_paths);
}

fn initializeCom() !ComInit {
    var state = ComInit{};

    const hr = c.CoInitializeEx(null, c.COINIT_APARTMENTTHREADED | c.COINIT_DISABLE_OLE1DDE);
    if (hr == c.RPC_E_CHANGED_MODE) {
        state.co_uninit = false;
    } else if (succeeded(hr)) {
        state.co_uninit = true;
    } else {
        return error.CoInitializeFailed;
    }

    const hr_ole = c.OleInitialize(null);
    if (hr_ole == c.RPC_E_CHANGED_MODE) {
        state.ole_uninit = false;
    } else if (succeeded(hr_ole)) {
        state.ole_uninit = true;
    } else {
        return error.OleInitializeFailed;
    }

    return state;
}

fn waitForClipboardUpdate(previous_seq: c.DWORD) void {
    var msg: c.MSG = undefined;
    var attempts: usize = 0;
    while (attempts < 80) : (attempts += 1) {
        while (c.PeekMessageW(&msg, null, 0, 0, c.PM_REMOVE) != 0) {
            _ = c.TranslateMessage(&msg);
            _ = c.DispatchMessageW(&msg);
        }

        const seq_now = c.GetClipboardSequenceNumber();
        if (seq_now != 0 and seq_now != previous_seq) return;

        c.Sleep(25);
    }
}

fn isLikelyCopyCommand(
    allocator: std.mem.Allocator,
    menu_vtbl: [*c]c.IContextMenuVtbl,
    context_menu: *c.IContextMenu,
    hmenu: c.HMENU,
    selected_cmd: c.UINT,
    verb_offset: usize,
) bool {
    if (menu_vtbl.*.GetCommandString) |get_command_string| {
        var verb_buf: [64]u8 = [_]u8{0} ** 64;
        const hr_verb = get_command_string(
            context_menu,
            @as(c.UINT_PTR, @intCast(verb_offset)),
            @as(c.UINT, @intCast(c.GCS_VERBA)),
            null,
            @ptrCast(&verb_buf[0]),
            @as(c.UINT, @intCast(verb_buf.len)),
        );
        if (succeeded(hr_verb)) {
            const verb_len = std.mem.indexOfScalar(u8, &verb_buf, 0) orelse verb_buf.len;
            const verb = verb_buf[0..verb_len];
            diagLog(allocator, "probe VERBA hr={d} value=\"{s}\"", .{ hr_verb, verb });
            if (asciiLooksLikeCopy(verb)) return true;
        } else {
            diagLog(allocator, "probe VERBA hr={d}", .{hr_verb});
        }

        var verb_w: [128]u16 = [_]u16{0} ** 128;
        const hr_verb_w = get_command_string(
            context_menu,
            @as(c.UINT_PTR, @intCast(verb_offset)),
            @as(c.UINT, @intCast(c.GCS_VERBW)),
            null,
            @ptrCast(&verb_w[0]),
            @as(c.UINT, @intCast(verb_w.len)),
        );
        if (succeeded(hr_verb_w)) {
            const len_w = std.mem.indexOfScalar(u16, &verb_w, 0) orelse verb_w.len;
            if (std.unicode.utf16LeToUtf8Alloc(allocator, verb_w[0..len_w])) |verb_utf8| {
                defer allocator.free(verb_utf8);
                diagLog(allocator, "probe VERBW hr={d} value=\"{s}\"", .{ hr_verb_w, verb_utf8 });
            } else |_| {
                diagLog(allocator, "probe VERBW hr={d} value=<utf16-convert-failed>", .{hr_verb_w});
            }
            if (utf16LooksLikeCopy(verb_w[0..len_w])) return true;
        } else {
            diagLog(allocator, "probe VERBW hr={d}", .{hr_verb_w});
        }

        var help_w: [256]u16 = [_]u16{0} ** 256;
        const hr_help_w = get_command_string(
            context_menu,
            @as(c.UINT_PTR, @intCast(verb_offset)),
            @as(c.UINT, @intCast(c.GCS_HELPTEXTW)),
            null,
            @ptrCast(&help_w[0]),
            @as(c.UINT, @intCast(help_w.len)),
        );
        if (succeeded(hr_help_w)) {
            const help_len = std.mem.indexOfScalar(u16, &help_w, 0) orelse help_w.len;
            if (std.unicode.utf16LeToUtf8Alloc(allocator, help_w[0..help_len])) |help_utf8| {
                defer allocator.free(help_utf8);
                diagLog(allocator, "probe HELPTEXTW hr={d} value=\"{s}\"", .{ hr_help_w, help_utf8 });
            } else |_| {
                diagLog(allocator, "probe HELPTEXTW hr={d} value=<utf16-convert-failed>", .{hr_help_w});
            }
            if (utf16LooksLikeCopy(help_w[0..help_len])) return true;
        } else {
            diagLog(allocator, "probe HELPTEXTW hr={d}", .{hr_help_w});
        }
    }

    const label_match = menuItemLabelLooksLikeCopy(allocator, hmenu, selected_cmd);
    if (label_match) {
        diagLog(allocator, "probe LABEL matched copy", .{});
    } else {
        diagLog(allocator, "probe LABEL no match", .{});
    }
    return label_match;
}

fn menuItemLabelLooksLikeCopy(allocator: std.mem.Allocator, hmenu: c.HMENU, selected_cmd: c.UINT) bool {
    var text_buf: [256]u16 = [_]u16{0} ** 256;
    const count = findMenuLabelRecursive(hmenu, selected_cmd, &text_buf) orelse return false;
    if (count == 0) return false;
    if (std.unicode.utf16LeToUtf8Alloc(allocator, text_buf[0..count])) |label_utf8| {
        defer allocator.free(label_utf8);
        diagLog(allocator, "probe LABEL value=\"{s}\"", .{label_utf8});
    } else |_| {
        diagLog(allocator, "probe LABEL value=<utf16-convert-failed>", .{});
    }
    return normalizedLabelStartsWithCopy(text_buf[0..count]);
}

fn findMenuLabelRecursive(hmenu: c.HMENU, selected_cmd: c.UINT, out: []u16) ?usize {
    const count = c.GetMenuItemCount(hmenu);
    if (count <= 0) return null;

    var pos: c_int = 0;
    while (pos < count) : (pos += 1) {
        const item_id = c.GetMenuItemID(hmenu, pos);
        if (item_id == selected_cmd) {
            const n = c.GetMenuStringW(
                hmenu,
                @as(c.UINT, @intCast(pos)),
                @ptrCast(out.ptr),
                @as(c_int, @intCast(out.len)),
                @as(c.UINT, @intCast(c.MF_BYPOSITION)),
            );
            if (n > 0) return @intCast(n);
            return 0;
        }

        const sub = c.GetSubMenu(hmenu, pos);
        if (sub != null) {
            if (findMenuLabelRecursive(sub, selected_cmd, out)) |len| {
                return len;
            }
        }
    }
    return null;
}

fn normalizedLabelStartsWithCopy(label: []const u16) bool {
    var ascii_buf: [128]u8 = undefined;
    var n: usize = 0;

    for (label) |ch| {
        if (ch == 0 or ch == '\t') break;
        if (ch == '&' or ch == ' ' or ch == ':') continue;
        if (ch > 0x7F) continue;
        if (n >= ascii_buf.len) break;
        ascii_buf[n] = std.ascii.toLower(@as(u8, @intCast(ch)));
        n += 1;
    }

    const s = ascii_buf[0..n];
    return asciiLooksLikeCopy(s);
}

fn asciiContainsNoCase(haystack: []const u8, needle: []const u8) bool {
    if (needle.len == 0) return true;
    if (needle.len > haystack.len) return false;

    var i: usize = 0;
    while (i + needle.len <= haystack.len) : (i += 1) {
        var ok = true;
        var j: usize = 0;
        while (j < needle.len) : (j += 1) {
            if (std.ascii.toLower(haystack[i + j]) != std.ascii.toLower(needle[j])) {
                ok = false;
                break;
            }
        }
        if (ok) return true;
    }
    return false;
}

fn asciiLooksLikeCopy(s: []const u8) bool {
    const looks_like_copy = asciiContainsNoCase(s, "copy") or asciiContainsNoCase(s, "kopier");
    const looks_like_copy_path = asciiContainsNoCase(s, "path") or asciiContainsNoCase(s, "pfad");
    return looks_like_copy and !looks_like_copy_path;
}

fn utf16ContainsNoCaseAscii(haystack: []const u16, needle: []const u8) bool {
    if (needle.len == 0) return true;
    if (needle.len > haystack.len) return false;

    var i: usize = 0;
    while (i + needle.len <= haystack.len) : (i += 1) {
        var ok = true;
        var j: usize = 0;
        while (j < needle.len) : (j += 1) {
            const ch = haystack[i + j];
            if (ch > 0x7F) {
                ok = false;
                break;
            }
            const lower_h = std.ascii.toLower(@as(u8, @intCast(ch)));
            const lower_n = std.ascii.toLower(needle[j]);
            if (lower_h != lower_n) {
                ok = false;
                break;
            }
        }
        if (ok) return true;
    }
    return false;
}

fn utf16LooksLikeCopy(s: []const u16) bool {
    const looks_like_copy = utf16ContainsNoCaseAscii(s, "copy") or utf16ContainsNoCaseAscii(s, "kopier");
    const looks_like_copy_path = utf16ContainsNoCaseAscii(s, "path") or utf16ContainsNoCaseAscii(s, "pfad");
    return looks_like_copy and !looks_like_copy_path;
}

fn setClipboardFileDropList(allocator: std.mem.Allocator, paths: []const []const u8) !void {
    if (paths.len == 0) return error.NoPathsForClipboard;
    diagLog(allocator, "setClipboardFileDropList begin paths={d}", .{paths.len});

    var wide_paths: std.ArrayList([:0]u16) = .empty;
    defer {
        for (wide_paths.items) |w| allocator.free(w);
        wide_paths.deinit(allocator);
    }

    var total_chars: usize = 1; // final extra NUL for double terminator
    for (paths) |path| {
        const windows_path = try normalizeWindowsPathAlloc(allocator, path);
        defer allocator.free(windows_path);
        const wide = try std.unicode.utf8ToUtf16LeAllocZ(allocator, windows_path);
        try wide_paths.append(allocator, wide);
        total_chars += wide.len + 1;
    }

    const total_bytes: usize = @sizeOf(c.DROPFILES) + total_chars * @sizeOf(u16);
    const hmem = c.GlobalAlloc(c.GMEM_MOVEABLE | c.GMEM_ZEROINIT, total_bytes);
    if (hmem == null) return error.GlobalAllocFailed;
    diagLog(allocator, "GlobalAlloc ok bytes={d}", .{total_bytes});

    var transfer_to_clipboard = false;
    defer {
        if (!transfer_to_clipboard) {
            _ = c.GlobalFree(hmem);
        }
    }

    const raw = c.GlobalLock(hmem);
    if (raw == null) return error.GlobalLockFailed;

    const drop: *c.DROPFILES = @ptrCast(@alignCast(raw));
    drop.* = std.mem.zeroes(c.DROPFILES);
    drop.pFiles = @as(c.DWORD, @intCast(@sizeOf(c.DROPFILES)));
    drop.fWide = 1;

    const base_bytes: [*]u8 = @ptrCast(raw);
    const names_ptr: [*]u16 = @ptrCast(@alignCast(base_bytes + @sizeOf(c.DROPFILES)));
    const names: []u16 = names_ptr[0..total_chars];

    var cursor: usize = 0;
    for (wide_paths.items) |w| {
        std.mem.copyForwards(u16, names[cursor .. cursor + w.len], w);
        cursor += w.len;
        names[cursor] = 0;
        cursor += 1;
    }
    names[cursor] = 0;

    _ = c.GlobalUnlock(hmem);

    if (!openClipboardWithRetry()) return error.OpenClipboardFailed;
    defer _ = c.CloseClipboard();
    diagLog(allocator, "OpenClipboard ok", .{});

    if (c.EmptyClipboard() == 0) return error.EmptyClipboardFailed;
    diagLog(allocator, "EmptyClipboard ok", .{});

    if (c.SetClipboardData(@as(c.UINT, @intCast(c.CF_HDROP)), hmem) == null) {
        return error.SetClipboardDataFailed;
    }
    diagLog(allocator, "SetClipboardData CF_HDROP ok", .{});

    // Hint shell paste targets that this is a copy operation, not move.
    const preferred_format = c.RegisterClipboardFormatW(std.unicode.utf8ToUtf16LeStringLiteral("Preferred DropEffect"));
    if (preferred_format != 0) {
        const effect_hmem = c.GlobalAlloc(c.GMEM_MOVEABLE | c.GMEM_ZEROINIT, @sizeOf(c.DWORD));
        if (effect_hmem != null) {
            var effect_transferred = false;
            defer {
                if (!effect_transferred) _ = c.GlobalFree(effect_hmem);
            }

            const effect_raw = c.GlobalLock(effect_hmem);
            if (effect_raw != null) {
                const effect_ptr: *c.DWORD = @ptrCast(@alignCast(effect_raw));
                effect_ptr.* = @as(c.DWORD, @intCast(c.DROPEFFECT_COPY));
                _ = c.GlobalUnlock(effect_hmem);

                if (c.SetClipboardData(preferred_format, effect_hmem) != null) {
                    effect_transferred = true;
                }
            }
        }
    }

    transfer_to_clipboard = true;
    _ = c.OleFlushClipboard();
    diagLog(allocator, "setClipboardFileDropList done", .{});
}

fn normalizeWindowsPathAlloc(allocator: std.mem.Allocator, path: []const u8) ![]u8 {
    const out = try allocator.dupe(u8, path);
    for (out) |*ch| {
        if (ch.* == '/') ch.* = '\\';
    }
    return out;
}

fn openClipboardWithRetry() bool {
    var attempts: usize = 0;
    while (attempts < 80) : (attempts += 1) {
        if (c.OpenClipboard(null) != 0) return true;
        c.Sleep(25);
    }
    return false;
}

fn normalizePaths(allocator: std.mem.Allocator, paths: []const []const u8) ![]const []const u8 {
    var list: std.ArrayList([]const u8) = .empty;
    errdefer {
        for (list.items) |p| allocator.free(p);
        list.deinit(allocator);
    }

    if (paths.len == 0) {
        const abs = try std.fs.cwd().realpathAlloc(allocator, ".");
        try list.append(allocator, abs);
    } else {
        for (paths) |path| {
            const abs = try std.fs.cwd().realpathAlloc(allocator, path);
            try list.append(allocator, abs);
        }
    }

    return list.toOwnedSlice(allocator);
}

fn freePathList(allocator: std.mem.Allocator, absolute_paths: []const []const u8) void {
    for (absolute_paths) |path| allocator.free(path);
    allocator.free(absolute_paths);
}

fn parseDisplayNameToPidl(allocator: std.mem.Allocator, path: []const u8) !*c.ITEMIDLIST {
    const path_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, path);
    defer allocator.free(path_w);

    var pidl: ?*c.ITEMIDLIST = null;
    var attributes: c.ULONG = 0;
    const hr = c.SHParseDisplayName(path_w.ptr, null, &pidl, 0, &attributes);
    if (!succeeded(hr) or pidl == null) {
        return error.ParseDisplayNameFailed;
    }
    return pidl.?;
}

fn acquireOwnerWindow(mode: WindowMode) !OwnerWindow {
    if (mode == .auto) {
        const foreground = c.GetForegroundWindow();
        if (foreground != null) {
            return .{ .hwnd = foreground, .owned = false };
        }
    }

    const hwnd = try createHiddenOwnerWindow();
    return .{
        .hwnd = hwnd,
        .owned = true,
    };
}

fn createHiddenOwnerWindow() !c.HWND {
    const instance = c.GetModuleHandleW(null);
    if (instance == null) return error.GetModuleHandleFailed;

    var wc: c.WNDCLASSW = std.mem.zeroes(c.WNDCLASSW);
    wc.lpfnWndProc = menuWndProc;
    wc.hInstance = instance;
    wc.lpszClassName = hidden_window_class;

    if (c.RegisterClassW(&wc) == 0) {
        const last_err = c.GetLastError();
        if (last_err != c.ERROR_CLASS_ALREADY_EXISTS) {
            return error.RegisterClassFailed;
        }
    }

    const hwnd = c.CreateWindowExW(
        0,
        hidden_window_class,
        hidden_window_class,
        0,
        0,
        0,
        0,
        0,
        null,
        null,
        instance,
        null,
    );
    if (hwnd == null) return error.CreateWindowFailed;
    return hwnd;
}

fn resolvePoint(x: ?i32, y: ?i32) !c.POINT {
    var point: c.POINT = undefined;
    if (c.GetCursorPos(&point) == 0) return error.GetCursorPosFailed;
    if (x) |xv| point.x = xv;
    if (y) |yv| point.y = yv;
    return point;
}

fn attachMenuMsgForwarders(context_menu: *c.IContextMenu) void {
    clearMenuMsgForwarders();

    const unk: [*c]c.IUnknown = @ptrCast(context_menu);
    const unk_vtbl = unk.*.lpVtbl orelse return;
    const query_interface = unk_vtbl.*.QueryInterface orelse return;

    var ppv3: ?*anyopaque = null;
    if (succeeded(query_interface(unk, &c.IID_IContextMenu3, &ppv3)) and ppv3 != null) {
        g_context_menu3 = @ptrCast(@alignCast(ppv3.?));
        return;
    }

    var ppv2: ?*anyopaque = null;
    if (succeeded(query_interface(unk, &c.IID_IContextMenu2, &ppv2)) and ppv2 != null) {
        g_context_menu2 = @ptrCast(@alignCast(ppv2.?));
    }
}

fn clearMenuMsgForwarders() void {
    if (g_context_menu3) |ctx3| {
        releaseIUnknown(ctx3);
        g_context_menu3 = null;
    }
    if (g_context_menu2) |ctx2| {
        releaseIUnknown(ctx2);
        g_context_menu2 = null;
    }
}

fn menuWndProc(hwnd: c.HWND, msg: c.UINT, w_param: c.WPARAM, l_param: c.LPARAM) callconv(.winapi) c.LRESULT {
    switch (msg) {
        c.WM_INITMENUPOPUP, c.WM_DRAWITEM, c.WM_MEASUREITEM, c.WM_MENUCHAR => {
            if (g_context_menu3) |ctx3| {
                var result: c.LRESULT = 0;
                const vtbl3 = ctx3.lpVtbl orelse return 0;
                if (vtbl3.*.HandleMenuMsg2) |handle_menu_msg2| {
                    const hr3 = handle_menu_msg2(ctx3, msg, w_param, l_param, &result);
                    if (succeeded(hr3)) return result;
                }
            }
            if (g_context_menu2) |ctx2| {
                const vtbl2 = ctx2.lpVtbl orelse return 0;
                if (vtbl2.*.HandleMenuMsg) |handle_menu_msg| {
                    const hr2 = handle_menu_msg(ctx2, msg, w_param, l_param);
                    if (succeeded(hr2)) return 0;
                }
            }
        },
        c.WM_COMMAND => {
            if (l_param == 0) {
                g_menu_command_id = @as(c.UINT, @intCast(w_param & 0xFFFF));
                return 0;
            }
        },
        else => {},
    }
    return c.DefWindowProcW(hwnd, msg, w_param, l_param);
}

fn installRegistryEntries(allocator: std.mem.Allocator) !void {
    const exe_path = try std.fs.selfExePathAlloc(allocator);
    defer allocator.free(exe_path);

    const command_for_item = try std.fmt.allocPrint(allocator, "\"{s}\" \"%1\"", .{exe_path});
    defer allocator.free(command_for_item);

    const command_for_background = try std.fmt.allocPrint(allocator, "\"{s}\" \"%V\"", .{exe_path});
    defer allocator.free(command_for_background);

    try writeMenuEntry(
        allocator,
        "Software\\Classes\\*\\shell\\winctxmenu-zig",
        "Show Explorer Context Menu",
        command_for_item,
        exe_path,
    );
    try writeMenuEntry(
        allocator,
        "Software\\Classes\\Directory\\shell\\winctxmenu-zig",
        "Show Explorer Context Menu",
        command_for_item,
        exe_path,
    );
    try writeMenuEntry(
        allocator,
        "Software\\Classes\\Directory\\Background\\shell\\winctxmenu-zig",
        "Show Explorer Context Menu",
        command_for_background,
        exe_path,
    );
}

fn uninstallRegistryEntries(allocator: std.mem.Allocator) !void {
    try deleteRegistryTree(allocator, "Software\\Classes\\*\\shell\\winctxmenu-zig");
    try deleteRegistryTree(allocator, "Software\\Classes\\Directory\\shell\\winctxmenu-zig");
    try deleteRegistryTree(allocator, "Software\\Classes\\Directory\\Background\\shell\\winctxmenu-zig");
}

fn writeMenuEntry(
    allocator: std.mem.Allocator,
    key_path_utf8: []const u8,
    display_utf8: []const u8,
    command_utf8: []const u8,
    icon_utf8: []const u8,
) !void {
    const command_key_utf8 = try std.fmt.allocPrint(allocator, "{s}\\command", .{key_path_utf8});
    defer allocator.free(command_key_utf8);

    const key_path_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, key_path_utf8);
    defer allocator.free(key_path_w);
    const command_key_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, command_key_utf8);
    defer allocator.free(command_key_w);
    const display_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, display_utf8);
    defer allocator.free(display_w);
    const command_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, command_utf8);
    defer allocator.free(command_w);
    const icon_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, icon_utf8);
    defer allocator.free(icon_w);
    const icon_name_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, "Icon");
    defer allocator.free(icon_name_w);

    const root_key = try regCreateKey(key_path_w);
    defer _ = RegCloseKey(root_key);
    try regSetSz(root_key, null, display_w);
    try regSetSz(root_key, icon_name_w.ptr, icon_w);

    const command_key = try regCreateKey(command_key_w);
    defer _ = RegCloseKey(command_key);
    try regSetSz(command_key, null, command_w);
}

fn deleteRegistryTree(allocator: std.mem.Allocator, key_path_utf8: []const u8) !void {
    const key_path_w = try std.unicode.utf8ToUtf16LeAllocZ(allocator, key_path_utf8);
    defer allocator.free(key_path_w);

    const status = RegDeleteTreeW(win.HKEY_CURRENT_USER, key_path_w.ptr);
    if (status == 0 or status == 2 or status == 3) {
        return;
    }
    return error.RegistryDeleteFailed;
}

fn regCreateKey(subkey: [:0]const u16) !win.HKEY {
    var key: win.HKEY = undefined;
    const status = RegCreateKeyExW(
        win.HKEY_CURRENT_USER,
        subkey.ptr,
        0,
        null,
        0,
        win.KEY_SET_VALUE | win.KEY_CREATE_SUB_KEY,
        null,
        &key,
        null,
    );
    if (status != 0) return error.RegistryCreateFailed;
    return key;
}

fn regSetSz(key: win.HKEY, value_name: ?[*:0]const u16, data: [:0]const u16) !void {
    const bytes: win.DWORD = @intCast((data.len + 1) * @sizeOf(u16));
    const status = RegSetValueExW(
        key,
        value_name,
        0,
        win.REG.SZ,
        @ptrCast(data.ptr),
        bytes,
    );
    if (status != 0) return error.RegistrySetValueFailed;
}

fn releaseCom(ptr: ?*anyopaque) void {
    if (ptr) |p| {
        releaseIUnknown(p);
    }
}

fn releaseIUnknown(ptr: anytype) void {
    const unk: *c.IUnknown = @ptrCast(@alignCast(ptr));
    const vtbl = unk.lpVtbl orelse return;
    const release_fn = vtbl.*.Release orelse return;
    _ = release_fn(unk);
}

fn succeeded(hr: c.HRESULT) bool {
    return hr >= 0;
}

fn diagLogPathAlloc(allocator: std.mem.Allocator) ![]u8 {
    if (std.process.getEnvVarOwned(allocator, "TEMP")) |temp| {
        defer allocator.free(temp);
        return std.fs.path.join(allocator, &.{ temp, "winctxmenu-debug.log" });
    } else |_| {}

    if (std.process.getEnvVarOwned(allocator, "TMP")) |tmp| {
        defer allocator.free(tmp);
        return std.fs.path.join(allocator, &.{ tmp, "winctxmenu-debug.log" });
    } else |_| {}

    return allocator.dupe(u8, "winctxmenu-debug.log");
}

fn diagLog(allocator: std.mem.Allocator, comptime fmt: []const u8, args: anytype) void {
    const msg = std.fmt.allocPrint(allocator, fmt, args) catch return;
    defer allocator.free(msg);

    const line = std.fmt.allocPrint(allocator, "[{d}] {s}\r\n", .{ std.time.milliTimestamp(), msg }) catch return;
    defer allocator.free(line);

    const path = diagLogPathAlloc(allocator) catch return;
    defer allocator.free(path);

    var file = if (std.fs.path.isAbsolute(path))
        std.fs.openFileAbsolute(path, .{ .mode = .read_write }) catch |err| switch (err) {
            error.FileNotFound => std.fs.createFileAbsolute(path, .{ .truncate = false, .read = true }) catch return,
            else => return,
        }
    else
        std.fs.cwd().openFile(path, .{ .mode = .read_write }) catch |err| switch (err) {
            error.FileNotFound => std.fs.cwd().createFile(path, .{ .truncate = false, .read = true }) catch return,
            else => return,
        };
    defer file.close();

    file.seekFromEnd(0) catch return;
    _ = file.writeAll(line) catch return;
}

fn printUsage() void {
    std.debug.print(
        \\winctxmenu - native Explorer context menu wrapper
        \\
        \\Usage:
        \\  winctxmenu [--x N] [--y N] [--window hidden|auto] [path ...]
        \\  winctxmenu --install
        \\  winctxmenu --uninstall
        \\  winctxmenu --copy-hdrop <path ...>
        \\
        \\Behavior:
        \\  - With one or more paths: opens native context menu for that selection.
        \\  - With no paths: opens context menu for current directory.
        \\  - All paths must share the same parent directory.
        \\
    , .{});
}
