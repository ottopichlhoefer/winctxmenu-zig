# winctxmenu

`winctxmenu` is a small Zig CLI that pops the **native Windows Explorer context menu** for an arbitrary file/folder selection from the command line.

It exists because Emacs Dired on Windows has no equivalent of *right-click → "Open with…"* / *"Git GUI Here"* / shell-extension entries from WinRAR, 7-Zig, TortoiseGit, WSL, Windows Terminal, etc. Rather than reimplement each integration, `winctxmenu` asks the real Shell to render its real menu against a real `IShellItemArray`, so you get exactly what Explorer would show you — including entries from any shell extension the user has installed.

---

## Background

### What it does under the hood

1. Parses the command-line selection and resolves each path to an `IShellItem` (via `SHCreateItemFromParsingName`).
2. Binds the selection's common parent folder as an `IShellFolder`.
3. Obtains an `IContextMenu` for the selection via `IShellFolder::GetUIObjectOf`.
4. Creates an invisible host window, populates a `CreatePopupMenu` with `QueryContextMenu(CMF_NORMAL | CMF_EXPLORE)`, and calls `TrackPopupMenuEx` at either the cursor or the supplied `--x`/`--y` screen coordinates.
5. Dispatches the selected verb back through `IContextMenu::InvokeCommand`.

All interop is direct COM — `advapi32`, `ole32`, `shell32`, `user32`, `uuid` linked statically via `build.zig`. No `.NET`, no PowerShell, no AutoHotkey.

### Why a helper process (and not an Emacs DLL)?

- **Isolation.** The Explorer context menu pump spins its own message loop; hosting it in-process with Emacs risks starving Emacs's own loop and has historically deadlocked on certain shell extensions (notably the OneDrive/SharePoint overlays).
- **Async.** Invoked via `start-process`, the menu is non-blocking. The user can dismiss it with Esc or pick a verb while Emacs keeps running.
- **No admin required.** Registry install mode uses `HKCU` keys only.

### How it's used from Emacs

`~/.doom.d/config.el` (`orp/show-win-context-menu`) binds `C-c c` in Dired: gathers `dired-get-marked-files` (or the file at point, or the current directory), shells out to `winctxmenu.exe --window hidden <paths>`, and returns. The menu appears at the cursor.

Path resolution goes PATH → local build, so the same `config.el` works on a machine with the Scoop-installed binary and on the dev machine that builds from source.

---

## Install

### Via Scoop (recommended)

No bucket required — point Scoop at the manifest in this repo:

```powershell
scoop install https://raw.githubusercontent.com/ottopichlhoefer/winctxmenu-zig/master/scoop/winctxmenu.json
```

Upgrades:

```powershell
scoop update winctxmenu
```

The manifest's `autoupdate` block follows GitHub Releases, and the hash is resolved from the `.sha256` sidecar at install time — so new versions flow automatically without manifest edits.

### Via GitHub Release

```powershell
# Latest release
gh release download -R ottopichlhoefer/winctxmenu-zig -p winctxmenu.exe -D $env:USERPROFILE\bin

# Pinned version
gh release download v0.1.0 -R ottopichlhoefer/winctxmenu-zig -p winctxmenu.exe -D $env:USERPROFILE\bin
```

Verify the download:

```powershell
$expected = (Invoke-WebRequest https://github.com/ottopichlhoefer/winctxmenu-zig/releases/latest/download/winctxmenu.exe.sha256).Content.Split(' ')[0]
$actual   = (Get-FileHash $env:USERPROFILE\bin\winctxmenu.exe -Algorithm SHA256).Hash.ToLower()
if ($expected -eq $actual) { "OK" } else { "HASH MISMATCH" }
```

### From source

Requires [Zig](https://ziglang.org/) 0.15.2+ (known-good: 0.15.2 on Windows 11 / MSVC linker or `zig cc`).

```powershell
git clone https://github.com/ottopichlhoefer/winctxmenu-zig
cd winctxmenu-zig
zig build -Doptimize=ReleaseSafe
# Binary at zig-out/bin/winctxmenu.exe
```

`ReleaseSafe` is what CI ships. For a smaller artifact at the cost of safety checks, use `-Doptimize=ReleaseFast` or `ReleaseSmall`.

---

## Usage

```powershell
winctxmenu.exe [--x N] [--y N] [--window hidden|auto] [path ...]
winctxmenu.exe --install
winctxmenu.exe --uninstall
winctxmenu.exe --help
```

| Flag | Meaning |
| --- | --- |
| `--x N` / `--y N` | Screen coordinates to anchor the menu. Default: current cursor position. |
| `--window hidden` | Use an invisible owner window (default — best for spawn-and-forget from Emacs). |
| `--window auto` | Use a normal (though still off-screen) owner window. |
| `--install` | Add shell extension entries under HKCU (see below). |
| `--uninstall` | Remove the HKCU entries added by `--install`. |
| `--help` | Print usage. |
| `path ...` | One or more filesystem paths. All paths must share a parent directory. |

Behavior:

- Zero paths → context menu for the current directory (`.`).
- One path → item-level menu for that file/folder.
- Multiple paths → multi-select menu (verbs operate on the whole selection). Paths must resolve to the same parent.

Exit code is `0` on successful display/dismissal, non-zero if COM initialization or binding failed. The tool does **not** wait for the invoked verb to finish.

### Emacs Dired integration

```elisp
(defvar win-context-menu-program nil
  "Override path; if nil, resolved via `executable-find' with local fallback.")

(defun orp/win-context-menu-program ()
  (or win-context-menu-program
      (executable-find "winctxmenu")
      (let ((local "C:/Users/orp77/winctxmenu-zig/zig-out/bin/winctxmenu.exe"))
        (and (file-exists-p local) local))))

(defun orp/show-win-context-menu ()
  "Show native Explorer context menu for marked files, else file at point."
  (interactive)
  (let* ((files (dired-get-marked-files nil nil nil t))
         (marked (cond
                  ((null (cdr files)) nil)
                  ((and (= (length files) 2) (eq (car files) t)) (list (cadr files)))
                  (t files)))
         (point-file (ignore-errors (dired-get-file-for-visit)))
         (targets (or marked
                      (and point-file (list point-file))
                      (list (dired-current-directory))))
         (prog (orp/win-context-menu-program)))
    (when prog
      (apply #'start-process "winctxmenu" "*WinCtxMenu*"
             prog "--window" "hidden" targets))))

(map! :map dired-mode-map "C-c c" #'orp/show-win-context-menu)
```

---

## Registry install mode

`--install` writes three HKCU shell entries that add a `winctxmenu-zig` verb to Explorer itself — useful if you want the tool accessible from Explorer's background context menu too:

- `HKCU\Software\Classes\*\shell\winctxmenu-zig` — file selections
- `HKCU\Software\Classes\Directory\shell\winctxmenu-zig` — folder selections
- `HKCU\Software\Classes\Directory\Background\shell\winctxmenu-zig` — empty Explorer window background

`--uninstall` removes exactly those three subtrees. No administrator privileges required — everything is user-scoped.

---

## Releasing (maintainer)

Tag and push:

```powershell
git tag v0.2.0
git push origin v0.2.0
```

The [`release`](.github/workflows/release.yml) workflow runs on `windows-latest`:

1. Installs Zig via `mlugg/setup-zig@v2`.
2. `zig build -Doptimize=ReleaseSafe`.
3. Computes a lowercase SHA-256 and writes `winctxmenu.exe.sha256` (`<hash>  winctxmenu.exe`).
4. Creates a GitHub Release with both files attached and auto-generated notes.

Once the release is live, `scoop update winctxmenu` picks up the new version on any machine.

---

## Roadmap

- Right-click invocation without rendering the menu (direct-to-verb mode for scripting).
- Accept a stdin-separated list of paths (beyond what `cmd.exe` argv quoting can express).
- Optional `.zip`/portable archive in releases alongside the raw `.exe`.

---

## License

MIT — see [LICENSE](LICENSE).
