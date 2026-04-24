# Save-As Export Macro – Installation Guide

## Files
| File | Purpose |
|------|---------|
| `SaveAs_Export.bas` | Main macro logic (module) |
| `ExportDialog.frm`  | Dialog box (UserForm) |
| `MacroLogger.bas`   | Reusable Excel/CSV logger – import into any VBA macro project |

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
| CheckBox      | `chkSTP`         | STEP AP203 (.step) → saved in 3D STEP FILE\ sub-folder | Inside `fraFormats` |
| Label         | `lblPreview`     | Output file name preview:             |                                |
| Label         | `lblPreviewVal`  | *(blank)*                             | Filled at runtime; WordWrap=True |
| CheckBox      | `chkEmail`       | Draft an e-mail to Debbie Decker for drawing transmittal? | Default unchecked |
| CommandButton | `btnOK`          | Export                                | Default = True                 |
| CommandButton | `btnCancel`      | Cancel                                | Cancel = True                  |

### 4b – Import MacroLogger (optional but recommended)
1. **File → Import File…**
2. Browse to `MacroLogger.bas` and click **Open**.
3. This adds a `MacroLogger` module that any other module in the same project can call.
   See the **MacroLogger** section at the bottom of this file for usage details.

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

---

## MacroLogger – reusable logger for any VBA macro

`MacroLogger.bas` can be imported into **any** SolidWorks (or Office) VBA project to give it an Excel log with a CSV fallback automatically.

### What it creates

| File | When used |
|------|-----------|
| `<logDir>\<logName>_Log.xlsx` | Primary log (always attempted first) |
| `<logDir>\<logName>_Overflow.csv` | Fallback when the .xlsx is open by another user |

### Sheet layout

| Row | Content |
|-----|---------|
| 1 | **Total Runs** \| `=COUNTA(A4:A1048576)` (live count) — bold |
| 2 | *(empty gap)* |
| 3 | **Date \| Time \| User \| …your columns…** — bold header |
| 4+ | One row per `WriteLog` call |

`Date`, `Time`, and `User` are written automatically on every call.

### Calling convention

```vba
' Minimal example – call this from your macro's main export/action sub:
MacroLogger.WriteLog _
    "Z:\DAG\SOLIDWORKS MACRO\My Macro\Log\", _   ' logDir  (trailing \ optional)
    "MyMacro", _                                   ' logName (becomes MyMacro_Log.xlsx)
    Array("Job Number", "Result", "Notes"), _      ' column headers
    Array(jobNum,       result,   notes)            ' matching values
```

The log directory is created automatically if it does not exist.  
Values are coerced to strings via `CStr()` – format them before passing if needed.

### Example: adding logging to a new macro

```vba
Sub RunMyMacro()
    ' ... do your work ...

    Dim jobNum  As String : jobNum  = "420788"
    Dim outcome As String : outcome = "OK"
    Dim skipped As Long   : skipped = 3

    MacroLogger.WriteLog _
        "Z:\DAG\SOLIDWORKS MACRO\My Macro\Log\", _
        "MyMacro", _
        Array("Job Number", "Outcome", "Skipped Sheets"), _
        Array(jobNum, outcome, skipped)
End Sub
```

Running this three times produces a sheet like:

| Total Runs | 3 | | |
|---|---|---|---|
| | | | |
| **Date** | **Time** | **User** | **Job Number** | **Outcome** | **Skipped Sheets** |
| 2026-04-24 | 09:15:02 | jsmith | 420788 | OK | 3 |
| 2026-04-24 | 10:30:44 | jsmith | 420789 | OK | 0 |
| 2026-04-24 | 14:05:11 | mwilson | 420790 | OK | 1 |
