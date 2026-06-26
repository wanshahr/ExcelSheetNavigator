# Excel Sheet Navigator

A professional, permanently installed **VSTO Excel add-in** that adds a dockable,
persistent navigation panel for moving between worksheets quickly — across **all**
open workbooks, without macros and without modifying any file.

It behaves like enterprise tools (e.g. Kutools): a vertical, searchable list of all
sheets in the active workbook, one click to activate a sheet, with the current sheet
highlighted.

---

## How to use

1. **Open Excel** — the **Sheet Navigator** panel appears docked on the right.
2. **Click a sheet name** (or press **Enter**) to jump straight to it.
3. **Search** — start typing part of a name to filter the list.
4. **Pin favorites** — click a row's star (★) to keep important sheets at the top.
5. **Compact rows** — click the **☰** button in the header to fit more sheets without scrolling.
6. **Dock** the panel Left / Right / Floating from the dropdown — your choice is saved.
7. **Show/hide** it anytime from **Home → Navigator → Sheet Navigator**.

---

## Download & install (one-click installer)

No Visual Studio required, and **no administrator rights**.

**Download the latest ZIP from the [Releases page](https://github.com/wanshahr/ExcelSheetNavigator/releases/latest)**
(direct link: [ExcelSheetNavigator.zip](https://github.com/wanshahr/ExcelSheetNavigator/releases/latest/download/ExcelSheetNavigator.zip)).
Then:

1. Right-click the ZIP → **Properties** → tick **Unblock** (if shown) → **OK**, then extract it.
2. Close Excel.
3. Double-click **`Install.bat`** (choose *More info → Run anyway* if SmartScreen appears).
4. Open Excel — the panel is there. To remove later, double-click **`Uninstall.bat`**.

The installer copies the add-in to `%LOCALAPPDATA%\ExcelSheetNavigator`, trusts the
certificate for that user only, and registers it under `HKCU`. If the panel does not
appear, install the free [VSTO runtime](https://aka.ms/vstoredist) and run
`Install.bat` again.

### Publish a new release (maintainers)

Build the self-contained package from source, then upload it as a GitHub Release so
others can download it:

```powershell
.\scripts\Build-Package.ps1                 # builds Release → dist\ExcelSheetNavigator.zip
gh release create v1.1.0 dist\ExcelSheetNavigator.zip --title "Excel Sheet Navigator v1.1.0" --notes "What changed…"
# …or replace the asset on an existing release:
gh release upload v1.0.0 dist\ExcelSheetNavigator.zip --clobber
```

> The package is signed with the bundled self-signed development certificate. For
> wide distribution, sign with an organization code-signing certificate so Windows
> trusts the publisher automatically (replace the `.pfx`/`.cer` and the thumbprint).

---

## Settings (persisted per user)

| Setting | Default | Meaning |
|---------|---------|---------|
| `DockPosition` | `Right` | Left / Right / Floating |
| `PanelWidth` | `220` | Docked width in pixels |
| `PaneVisible` | `True` | Whether the panel is shown |
| `Compact` | `False` | Compact (dense) vs comfortable row height |
| `Favorites` | _(empty)_ | Newline-separated list of pinned sheet names |

Stored via `Properties.Settings.Default` (user scope) and restored on startup.

---

## Features

- **Vertical sheet list** showing every worksheet in the active workbook.
- **Click (or Enter) to activate** a worksheet instantly.
- **Case-insensitive search filter** to jump to a sheet by typing part of its name.
- **Refresh** button to re-read the workbook structure on demand.
- **Favorites** — click a sheet's star to pin it to the top; saved between sessions.
- **Compact / comfortable rows** — a header toggle (☰) switches row density so you can
  fit more sheets without scrolling; the choice is **saved** between sessions.
- **Dockable** Left / Right / Floating — the choice is **saved** between sessions.
- **Active-sheet highlight** plus hover feedback.
- **Long names truncate** with an ellipsis and show the full name in a tooltip.
- **Hidden sheets** are shown in grey italic and are never force-unhidden.
- Scales to **100+ sheets** with a smooth, owner-drawn list.
- **Read-only navigation** — never reads or writes cell data.
- **Ribbon toggle** (Home → Navigator → *Sheet Navigator*) to show/hide and reopen
  the panel after closing it.

---

## Architecture

| Layer | File | Responsibility |
|-------|------|----------------|
| UI | [SheetNavigatorControl.cs](ExcelSheetNavigator/Controls/SheetNavigatorControl.cs) | The dockable panel: list rendering, search, favorites, density toggle, dock selector, events. No Excel code. |
| Business logic | [ExcelService.cs](ExcelSheetNavigator/Services/ExcelService.cs) | All Excel interop: enumerate sheets, activate a sheet, read names. Releases COM objects. |
| Data model | [WorksheetInfo.cs](ExcelSheetNavigator/Models/WorksheetInfo.cs) | Plain DTO: name, index, active/visible flags. |
| Entry point | [ThisAddIn.cs](ExcelSheetNavigator/ThisAddIn.cs) | Lifecycle, per-window task panes, settings, Excel event wiring. |
| Ribbon | [NavigatorRibbon.cs](ExcelSheetNavigator/Ribbon/NavigatorRibbon.cs) / [NavigatorRibbon.xml](ExcelSheetNavigator/Ribbon/NavigatorRibbon.xml) | Toggle button on the Home tab. |

The panel raises events (activate / refresh / dock-change); `ThisAddIn` translates
them into Excel actions via `ExcelService`. Excel application events
(`WorkbookOpen`, `NewWorkbook`, `SheetActivate`, `SheetBeforeDelete`, plus
`WindowActivate` for modern single-document-interface Excel) keep the panel in sync.
Because Excel 2013+ uses one window per workbook, a separate task pane is created per
window so the navigator is present in **every** open file.

---

## Requirements

- Windows with **Microsoft Excel** (desktop, 2013 or later recommended).
- **Visual Studio 2022/2026** with the *Office/SharePoint development* workload, **or**
  the matching MSBuild + VSTO build targets.
- **.NET Framework 4.8**.
- **Visual Studio 2010 Tools for Office Runtime** (installed with the Office workload).

---

## Build

From the repository root:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe' `
    ExcelSheetNavigator.sln /p:Configuration=Debug
```

Use `/p:Configuration=Release` for a release build. Output is written to
`ExcelSheetNavigator\bin\<Configuration>\`, including the signed
`ExcelSheetNavigator.vsto` deployment manifest and `ExcelSheetNavigator.dll.manifest`.

### Code-signing certificate

VSTO **requires signed** ClickOnce manifests. This project signs with a self-signed
development certificate, `ExcelSheetNavigator/ExcelSheetNavigator_TemporaryKey.pfx`
(thumbprint `FFE84A0F4BF082EB0D87E9A64AF121FB474BB7D4`). The public part is at
[scripts/ExcelSheetNavigator.cer](scripts/ExcelSheetNavigator.cer).

> **The private `.pfx` key is not committed** (it is git-ignored). A fresh clone must
> supply its own signing key — generate a new self-signed certificate (or use your
> organization's) and update `ManifestKeyFile` / `ManifestCertificateThumbprint` in
> the `.csproj`.

---

## Install (current user, no admin)

```powershell
.\scripts\Register-AddIn.ps1                 # registers the Debug build
.\scripts\Register-AddIn.ps1 -Configuration Release
```

The script trusts the development certificate (current-user Trusted Publishers / Root)
and writes the per-user registration under
`HKCU\Software\Microsoft\Office\Excel\Addins\ExcelSheetNavigator` with a
`...\ExcelSheetNavigator.vsto|vstolocal` manifest. **Restart Excel** afterwards.

## Uninstall

```powershell
.\scripts\Unregister-AddIn.ps1
```

Removes the registration and the development certificate from the trust stores
(pass `-KeepCertificate` to leave the certificate in place). Restart Excel.
