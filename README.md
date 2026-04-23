# winctxmenu

`winctxmenu` is a Zig CLI wrapper that shows the native Windows Explorer context menu for file/folder selections.

It is designed for asynchronous invocation from Emacs Dired.

## Install

### Via Scoop (recommended)

Install the manifest directly without adding a bucket:

```powershell
scoop install https://raw.githubusercontent.com/ottopichlhoefer/winctxmenu-zig/master/scoop/winctxmenu.json
```

Upgrades via `scoop update winctxmenu` — the manifest auto-follows the latest GitHub Release.

### Via GitHub Release

```powershell
gh release download -R ottopichlhoefer/winctxmenu-zig -p winctxmenu.exe -D $env:USERPROFILE\bin
```

### From source

Requires Zig 0.15.2+.

```powershell
zig build -Doptimize=ReleaseSafe
# Binary at zig-out/bin/winctxmenu.exe
```

## Usage

```powershell
winctxmenu.exe [--x N] [--y N] [--window hidden|auto] [path ...]
winctxmenu.exe --install
winctxmenu.exe --uninstall
```

Behavior:

- If one or more paths are provided, it opens a native context menu for that selection.
- If no paths are provided, it defaults to the current directory (`.`).
- Paths must resolve to the same parent directory.

## Registry install mode

`--install` writes HKCU shell entries:

- `HKCU\Software\Classes\*\shell\winctxmenu-zig`
- `HKCU\Software\Classes\Directory\shell\winctxmenu-zig`
- `HKCU\Software\Classes\Directory\Background\shell\winctxmenu-zig`

`--uninstall` removes the same entries.

No administrator privileges are required because the tool uses user-scope (`HKCU`) keys.

## Releasing

Tag a commit; the `release` workflow builds `winctxmenu.exe` + `winctxmenu.exe.sha256` on `windows-latest` and attaches them to a GitHub Release.

```powershell
git tag v0.1.0
git push origin v0.1.0
```

The Scoop manifest's `autoupdate` block picks up the new version on `scoop update winctxmenu`.

## License

MIT — see [LICENSE](LICENSE).
