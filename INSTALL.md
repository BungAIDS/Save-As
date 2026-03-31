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

| Control type  | (Name)           | Caption / Text                        | Notes                          |
|---------------|------------------|---------------------------------------|--------------------------------|
| Label         | `lblTitle`       | Save-As Export Utility                | Bold, size 12                  |
| Label         | `lblJobType`     | Job Type:                             |                                |
| Label         | `lblJobTypeVal`  | *(blank)*                             | Filled at runtime              |
| Label         | `lblFolder`      | Save folder:                          |                                |
| Label         | `lblFolderVal`   | *(blank)*                             | Filled at runtime; WordWrap=True |
| Label         | `lblDrawing`     | Drawing / Job #:                      |                                |
| Label         | `lblDrawingName` | *(blank)*                             | Filled at runtime              |
| Label         | `lblRevision`    | Revision Letter:                      |                                |
| TextBox       | `txtRevision`    | *(blank)*                             | MaxLength = 2                  |
| Frame         | `fraFormats`     | Export Formats                        |                                |
| CheckBox      | `chkPDF`         | PDF (.pdf)                            | Inside `fraFormats`            |
| CheckBox      | `chkDWG`         | AutoCAD DWG (.dwg)                    | Inside `fraFormats`            |
| CheckBox      | `chkDXF`         | DXF (.dxf) → saved in DXF\ sub-folder | Inside `fraFormats`           |
| Label         | `lblPreview`     | Output file name preview:             |                                |
| Label         | `lblPreviewVal`  | *(blank)*                             | Filled at runtime; WordWrap=True |
| CommandButton | `btnOK`          | Export                                | Default = True                 |
| CommandButton | `btnCancel`      | Cancel                                | Cancel = True                  |

### 5 – Run the macro
1. Open (or activate) a SolidWorks drawing (`.SLDDRW`).
2. **Tools → Macros → Run…** → select `SaveAs_Export.swp`.
3. The **Save-As Export** dialog will appear.

---

## Folder Architecture

```
Z:\Solidworks\Current\JOBS\
├── GENERAL LINE\
│   └── 121-125\           ← range folder (first 3 digits of job#, groups of 5)
│       └── 12345\
│           ├── 12345.SLDDRW
│           ├── 12345-RevA.pdf
│           ├── 12345-RevA.dwg
│           ├── DXF\
│           │   └── 12345-RevA.dxf
│           └── History\
│               └── (archived old revisions)
├── HD-PFD\
│   └── 40XXXX\            ← first 2 digits + XXXX  (e.g. 401234 → 40XXXX)
│       └── 401234\
│           └── ...
└── HDX\
    └── 121-125\           ← same range formula as GENERAL LINE
        └── 12345\
            └── ...
```

### Range folder formula
Groups job numbers into bands of 5 based on the **first 3 digits**:
- `ceil(prefix / 5)` → n; start = `5*(n-1)+1`, end = `5*n`
- Example: job `12345` → prefix `123` → n=25 → folder `121-125`

### Revision archiving
Before writing new files the macro scans the drawing folder (and `DXF\`) for any file matching `<DrawingName>-Rev*.pdf/.dwg/.dxf` that is **not** the current revision and moves them to `History\`. Timestamp-suffix prevents collisions in `History\`.

### DXF folder
`DXF\` is created automatically if it does not exist.

### Path validation
If the drawing is not inside `Z:\Solidworks\Current\JOBS` or is not under a recognised job-type folder, the macro will warn and ask whether to continue.
