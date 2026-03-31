# Save-As Export Macro – Installation Guide

## Files
| File | Purpose |
|------|---------|
| `SaveAs_Export.bas` | Main macro logic (module) |
| `ExportDialog.frm`  | Dialog box (UserForm) |

---

## Step-by-Step Install in SolidWorks 2025

### 1 – Open the VBA Editor
1. In SolidWorks go to **Tools → Macros → Edit…**
2. In the "Open" dialog, type a new filename (e.g. `SaveAs_Export.swp`) and click **Save**. This creates a blank macro project and opens the VBA IDE.

### 2 – Import the module
1. In the VBA IDE menu: **File → Import File…**
2. Browse to `SaveAs_Export.bas` and click **Open**.

### 3 – Import the UserForm
1. **File → Import File…**
2. Browse to `ExportDialog.frm` and click **Open**.
   > **Note:** The `.frm` file must be accompanied by its `.frx` binary. Because `.frx` files cannot be represented as plain text, you will need to build the form controls manually (step 4) or use the pre-built `.swp` file when one is provided.

### 4 – Build the UserForm controls (if importing from source)
Open `ExportDialog` in the designer and add the following controls with the exact `(Name)` values shown:

| Control type | (Name)        | Caption / Text                  | Notes                          |
|--------------|---------------|---------------------------------|--------------------------------|
| Label        | `lblTitle`    | Save-As Export Utility          | Bold, size 12                  |
| Label        | `lblDrawing`  | Drawing:                        |                                |
| Label        | `lblDrawingName` | *(blank)*                    | Filled at runtime              |
| Label        | `lblRevision` | Revision Letter:                |                                |
| TextBox      | `txtRevision` | *(blank)*                       | MaxLength = 2                  |
| Frame        | `fraFormats`  | Export Formats                  |                                |
| CheckBox     | `chkPDF`      | PDF (.pdf)                      | Inside `fraFormats`            |
| CheckBox     | `chkDWG`      | AutoCAD DWG (.dwg)              | Inside `fraFormats`            |
| CheckBox     | `chkDXF`      | DXF (.dxf)                      | Inside `fraFormats`            |
| Label        | `lblPreview`  | Output file name preview:       |                                |
| Label        | `lblPreviewVal` | *(blank)*                     | Filled at runtime; WordWrap=True |
| CommandButton | `btnOK`      | Export                          | Default = True                 |
| CommandButton | `btnCancel`  | Cancel                          | Cancel = True                  |

### 5 – Run the macro
1. Open (or activate) a SolidWorks drawing (`.SLDDRW`).
2. **Tools → Macros → Run…** → select `SaveAs_Export.swp`.
3. The **Save-As Export** dialog will appear.

---

## How It Works

```
Drawing folder\
├── 12345.SLDDRW              ← your drawing
├── 12345-RevA.pdf            ← PDF export (same folder)
├── 12345-RevA.dwg            ← DWG export (same folder)
├── DXF\
│   └── 12345-RevA.dxf        ← DXF export (dedicated sub-folder)
└── History\
    ├── 12345-RevA.pdf        ← archived when Rev B is exported
    └── 12345-RevA.dwg
```

### Revision archiving
Before writing new files the macro scans the drawing folder (and the `DXF\` sub-folder) for any file matching `<DrawingName>-Rev*.pdf/.dwg/.dxf` that is **not** the current revision, and moves them to `History\`. If a file with the same name already exists in `History\` it is renamed with a timestamp (`_YYYYMMDD_HHmmss`) so nothing is lost.

### DXF folder
If the `DXF\` sub-folder does not exist, it is created automatically before saving the DXF.

---

## Folder Architecture (TBD)
The current baseline saves all files relative to the drawing's own folder. Once you provide your folder architecture, the path logic in `SaveAs_Export.bas` (`drawingFolder` variable and the functions that follow) will be updated to route files to the correct locations automatically.
