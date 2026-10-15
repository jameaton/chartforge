# Installing ChartForge

ChartForge is a single file, `ChartForge.ppam`. You load it once through PowerPoint's Add-ins menu, and PowerPoint loads it again every time it starts.

## 1. Download

Download **`ChartForge.ppam`** from the [latest release](../../../releases/latest). Keep it somewhere it won't be deleted, e.g. a `ChartForge` folder in Documents. PowerPoint loads it from that location each time.

## 2. Allow add-in macros

ChartForge is written in VBA, so PowerPoint must be allowed to ask before running macros.

- **Mac:** **PowerPoint → Settings (Preferences) → Security & Privacy → Macro Security**. Choose **Disable all macros with notification**.
- **Windows:** **File → Options → Trust Center → Trust Center Settings… → Macro Settings**. Choose **Disable VBA macros with notification**.

> If your organisation manages Office settings centrally and macros are blocked, ask IT to allow `ChartForge.ppam`.

## 3. Load the add-in

### Mac

1. Open PowerPoint and any presentation (a blank one is fine).
2. **Tools → PowerPoint Add-ins…**
3. Click **+**, choose `ChartForge.ppam`, and click **Open**.
4. Click **Enable Macros** when asked.
5. Make sure **ChartForge** is ticked and click **OK**.

### Windows

1. Open PowerPoint and any presentation.
2. **File → Options → Add-ins.**
3. At the bottom, set **Manage** to **PowerPoint Add-ins** and click **Go…**
4. Click **Add New…**, choose `ChartForge.ppam`, and click **Open**.
5. Click **Enable Macros** when asked, then **Close**.

A **ChartForge** tab now appears on the ribbon.

## Updating

1. **Tools → PowerPoint Add-ins…** (Mac) or **File → Options → Add-ins → Go…** (Windows).
2. Select **ChartForge** and click **−** / **Remove**.
3. Quit PowerPoint completely (**⌘Q** on Mac) and reopen it.
4. Load the new `ChartForge.ppam` as above.

Charts made with an earlier version stay editable.

## Uninstalling

Remove **ChartForge** from the Add-ins list as above, then delete the `.ppam` file. Existing charts remain in your decks as ordinary shapes.

## Troubleshooting

| Problem | Fix |
|---|---|
| No ChartForge tab after loading | Quit PowerPoint completely and reopen it. Check **ChartForge** is ticked in the Add-ins list. |
| "Macros have been disabled" | Set macro security as in step 2, then remove and re-add the add-in. |
| Buttons do nothing | Close any `.pptm` file that contains the ChartForge code. Its macros can clash with the add-in's. |
| Text is in an unexpected font | Install IBM Plex Sans, or click **Switch font** to use Calibri. |
| "Select a ChartForge chart first" | Click the chart once so the whole group is selected, not a shape inside it. |
