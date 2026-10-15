# Building ChartForge from source

PowerPoint can't load VBA from a text file directly, and it has no UI for custom ribbons. Building is therefore two steps: PowerPoint packages the code, then a script adds the ribbon and turns the package into an add-in.

You need PowerPoint (Mac or Windows) and Python 3 (standard library only).

## 1. Package the code in PowerPoint

1. Remove any loaded ChartForge add-in (**Tools → PowerPoint Add-ins…**, then **−**) and restart PowerPoint. This leaves only one VBA project, so the import can't go into the wrong one.
2. Create a blank presentation.
3. Open the editor: **Tools → Macro → Visual Basic Editor** (Mac) or **Alt+F11** (Windows).
4. In the **Project** panel, click **VBAProject (Presentation1)**.
5. **File → Import File…** and choose `src/ChartForge.bas`. A **ChartForge** module appears under *Modules*.
6. **Debug → Compile VBAProject.** It should finish without an error.
7. Close the editor. **File → Save As…**, choose **PowerPoint Macro-Enabled Presentation (.pptm)** and save as `ChartForge.pptm`.

## 2. Build the add-in

```bash
python3 tools/build_addin.py path/to/ChartForge.pptm
```

This writes `ChartForge.ppam` next to the `.pptm`. The script:

- adds `ribbon/customUI.xml` as the package's ribbon definition, and
- switches the package's main content type from *macro-enabled presentation* to *add-in*.

Load the result as described in [INSTALL.md](INSTALL.md).

## Making changes

- **Chart behaviour:** edit `src/ChartForge.bas`. Layouts are in `LayoutColumns`, `LayoutWaterfall`, `LayoutMekko` and `LayoutGantt`; colours and the default font are constants at the top.
- **Ribbon buttons:** edit `ribbon/customUI.xml`. Each button's `onAction` must match a `Public Sub AF_…(control As IRibbonControl)` in the module.
- Keep `src/ChartForge.bas` ASCII-only with CRLF line endings (`.gitattributes` enforces CRLF); the VBA importer is picky about both.

After changing the code, rebuild from step 1. After changing only the ribbon, rerun step 2 on your existing `.pptm`.
