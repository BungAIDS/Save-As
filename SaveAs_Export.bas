'==============================================================================
' SaveAs_Export.bas
' SolidWorks 2025 VBA Macro
'
' PURPOSE: Export an open .SLDDRW to PDF, DWG, and/or DXF with a revision
'          suffix.  Old revisions are automatically archived to a "History"
'          sub-folder.  DXF files are saved to a dedicated "DXF" sub-folder
'          (created automatically if absent).
'
' FOLDER ARCHITECTURE (from Z:\Solidworks\Current\JOBS):
'
'   GENERAL LINE  ...\GENERAL LINE\<range>\<JobNum>\   range = e.g. "121-125"
'   HD-PFD        ...\HD-PFD\<2digitXXXX>\<JobNum>\    e.g. "40XXXX"
'   HDX           ...\HDX\<range>\<JobNum>\             same range formula
'
'   Range formula: groups of 5 on the first 3 digits of the job number.
'     e.g. job 12345 → first 3 digits = 123 → ceil(123/5)=25 → 121-125
'
' OUTPUT inside the job folder:
'   <JobNum>-Rev<X>.pdf          (PDF, same folder as drawing)
'   <JobNum>-Rev<X>.dwg          (DWG, same folder as drawing)
'   DXF\<JobNum>-Rev<X>.dxf      (DXF, DXF sub-folder – auto-created)
'   History\<old files>          (archived prior revisions)
'
' USAGE:   Open or activate the drawing, then Tools > Macros > Run.
'==============================================================================
Option Explicit

' Root of all SolidWorks job drawings on the network
Private Const SW_ROOT As String = "Z:\Solidworks\Current\JOBS"

' Recognised job-type sub-folders (must match folder names exactly, upper case)
Private Const JOBTYPE_GENLINE  As String = "GENERAL LINE"
Private Const JOBTYPE_HDPFD    As String = "HD-PFD"
Private Const JOBTYPE_HDX      As String = "HDX"

'==============================================================================
' ENTRY POINT
'==============================================================================
Sub main()

    Dim swApp   As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw  As SldWorks.DrawingDoc

    Set swApp = Application.SldWorks

    '--- Validate active document ---
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "No document is open.  Please open a SolidWorks drawing (.SLDDRW) first.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    If swModel.GetType <> swDocDRAWING Then
        MsgBox "The active document is not a drawing (.SLDDRW)." & vbCrLf & _
               "Please activate a drawing and run the macro again.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    Set swDraw = swModel

    Dim drawingPath As String
    drawingPath = swModel.GetPathName

    If drawingPath = "" Then
        MsgBox "The drawing has never been saved.  Please save it first.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    '--- Parse folder and base name ---
    Dim drawingFolder   As String   ' e.g. Z:\Solidworks\Current\JOBS\GENERAL LINE\121-125\12345\
    Dim drawingBaseName As String   ' e.g. 12345

    drawingFolder   = Left(drawingPath, InStrRev(drawingPath, "\"))
    drawingBaseName = Mid(drawingPath, Len(drawingFolder) + 1)
    drawingBaseName = Left(drawingBaseName, InStrRev(drawingBaseName, ".") - 1)

    '--- Validate the drawing is inside the expected SW root ---
    Dim jobType As String
    jobType = DetectJobType(drawingFolder)

    If Not PathStartsWith(drawingFolder, SW_ROOT) Then
        Dim resp As Integer
        resp = MsgBox("This drawing is not inside the expected SolidWorks jobs folder:" & vbCrLf & _
                      "  " & SW_ROOT & vbCrLf & vbCrLf & _
                      "Current location: " & drawingFolder & vbCrLf & vbCrLf & _
                      "Continue anyway?", _
                      vbExclamation + vbYesNo, "Save-As Export – Path Warning")
        If resp = vbNo Then Exit Sub
        jobType = "(unknown)"
    ElseIf jobType = "" Then
        resp = MsgBox("Drawing is inside the SolidWorks root but is not under a recognised" & vbCrLf & _
                      "job-type folder (GENERAL LINE, HD-PFD, HDX)." & vbCrLf & vbCrLf & _
                      "Continue anyway?", _
                      vbExclamation + vbYesNo, "Save-As Export – Path Warning")
        If resp = vbNo Then Exit Sub
        jobType = "(unknown)"
    End If

    '--- Show export dialog ---
    Dim dlg As New ExportDialog
    dlg.DrawingBaseName = drawingBaseName
    dlg.JobType         = jobType
    dlg.DrawingFolder   = drawingFolder
    dlg.Show

    If dlg.Cancelled Then Exit Sub

    '--- Collect choices ---
    Dim revLetter As String
    Dim doPDF     As Boolean
    Dim doDWG     As Boolean
    Dim doDXF     As Boolean

    revLetter = UCase(Trim(dlg.RevisionLetter))
    doPDF     = dlg.ExportPDF
    doDWG     = dlg.ExportDWG
    doDXF     = dlg.ExportDXF

    If revLetter = "" Then
        MsgBox "No revision letter entered.  Export cancelled.", vbExclamation, "Save-As Export"
        Exit Sub
    End If

    If Not doPDF And Not doDWG And Not doDXF Then
        MsgBox "No export format selected.  Please check at least one box.", vbExclamation, "Save-As Export"
        Exit Sub
    End If

    '--- Build output root name  e.g. "12345-RevA" ---
    Dim exportRoot As String
    exportRoot = drawingBaseName & "-Rev" & revLetter

    '--- Archive any previous revisions ---
    ArchiveOldRevisions drawingFolder, drawingBaseName, exportRoot

    '--- Export ---
    Dim errors   As Long
    Dim warnings As Long
    Dim outPath  As String
    Dim ok       As Boolean
    Dim results  As String
    results = ""

    If doPDF Then
        outPath = drawingFolder & exportRoot & ".pdf"
        ok = ExportToPDF(swDraw, outPath, errors, warnings)
        If ok Then
            results = results & "  PDF: " & outPath & vbCrLf
        Else
            MsgBox "PDF export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If doDWG Then
        outPath = drawingFolder & exportRoot & ".dwg"
        ok = ExportToDWG(swDraw, outPath, errors, warnings)
        If ok Then
            results = results & "  DWG: " & outPath & vbCrLf
        Else
            MsgBox "DWG export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If doDXF Then
        Dim dxfFolder As String
        dxfFolder = EnsureDXFFolder(drawingFolder)
        outPath = dxfFolder & exportRoot & ".dxf"
        ok = ExportToDXF(swDraw, outPath, errors, warnings)
        If ok Then
            results = results & "  DXF: " & outPath & vbCrLf
        Else
            MsgBox "DXF export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If results <> "" Then
        MsgBox "Export complete!" & vbCrLf & vbCrLf & results, vbInformation, "Save-As Export"
    End If

End Sub

'==============================================================================
' PATH HELPERS
'==============================================================================

' True if pathToCheck begins with rootPath (case-insensitive, tolerates
' trailing backslash differences).
Private Function PathStartsWith(ByVal pathToCheck As String, _
                                ByVal rootPath As String) As Boolean
    Dim a As String : a = LCase(NormalizeTrailingSlash(pathToCheck))
    Dim b As String : b = LCase(NormalizeTrailingSlash(rootPath))
    PathStartsWith = (Left(a, Len(b)) = b)
End Function

Private Function NormalizeTrailingSlash(ByVal p As String) As String
    If Right(p, 1) <> "\" Then p = p & "\"
    NormalizeTrailingSlash = p
End Function

'==============================================================================
' JOB TYPE DETECTION
' Inspects the drawing folder path for the recognised job-type sub-folder names.
' Returns one of JOBTYPE_* constants, or "" if not found.
'==============================================================================
Private Function DetectJobType(ByVal folderPath As String) As String
    Dim p As String
    p = LCase(folderPath)

    If InStr(p, "\" & LCase(JOBTYPE_GENLINE) & "\") > 0 Then
        DetectJobType = JOBTYPE_GENLINE
    ElseIf InStr(p, "\" & LCase(JOBTYPE_HDPFD) & "\") > 0 Then
        DetectJobType = JOBTYPE_HDPFD
    ElseIf InStr(p, "\" & LCase(JOBTYPE_HDX) & "\") > 0 Then
        DetectJobType = JOBTYPE_HDX
    Else
        DetectJobType = ""
    End If
End Function

'==============================================================================
' RANGE FOLDER CALCULATION
' Mirrors the PowerShell Calculate-Range function.
' Groups job numbers into bands of 5 based on the first 3 digits.
'   e.g. prefix=123 → n=ceil(123/5)=25 → start=121, end=125 → "121-125"
'   Special case: start=401,end=405 → start=400  (legacy band)
'==============================================================================
Public Function CalculateRange(ByVal prefix3 As Long) As String
    Dim n     As Long
    Dim start As Long
    Dim finish As Long

    n      = CLng(Int((prefix3 + 4) / 5))   ' equivalent of Ceiling(prefix3/5)
    start  = 5 * (n - 1) + 1
    finish = 5 * n

    ' Legacy special case preserved from PowerShell script
    If start = 401 And finish = 405 Then start = 400

    CalculateRange = CStr(start) & "-" & CStr(finish)
End Function

'==============================================================================
' INTERMEDIATE FOLDER – derive the intermediate folder name from a job number
' string and job type.  Mirrors the PowerShell Build-OtherJobFolderPath logic
' for the SolidWorks side.
'
' Returns "" on error (e.g. non-numeric prefix).
'==============================================================================
Public Function IntermediateFolder(ByVal jobNum As String, _
                                   ByVal jobType As String) As String
    Dim prefix3Int As Long
    Dim prefix2    As String
    Dim ok         As Boolean

    Select Case UCase(jobType)

        Case JOBTYPE_GENLINE, JOBTYPE_HDX
            ' First 3 digits → range folder
            If Len(jobNum) < 3 Then
                IntermediateFolder = ""
                Exit Function
            End If
            ok = IsNumeric(Left(jobNum, 3))
            If Not ok Then
                IntermediateFolder = ""
                Exit Function
            End If
            prefix3Int = CLng(Left(jobNum, 3))
            IntermediateFolder = CalculateRange(prefix3Int)

        Case JOBTYPE_HDPFD
            ' First 2 digits → e.g. "40XXXX"
            If Len(jobNum) < 2 Then
                IntermediateFolder = ""
                Exit Function
            End If
            prefix2 = Left(jobNum, 2)
            If Not IsNumeric(prefix2) Then
                IntermediateFolder = ""
                Exit Function
            End If
            IntermediateFolder = prefix2 & "XXXX"

        Case Else
            IntermediateFolder = ""

    End Select
End Function

'==============================================================================
' ARCHIVE - move all previous revision files to <drawingFolder>\History\
'           Scans both the drawing folder and the DXF sub-folder.
'           Files matching <baseName>-Rev*.ext that are NOT the current
'           revision are moved.  Timestamp-suffix prevents collisions.
'==============================================================================
Private Sub ArchiveOldRevisions(ByVal folder As String, _
                                ByVal baseName As String, _
                                ByVal currentRoot As String)

    Dim histFolder As String
    histFolder = folder & "History\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Extensions to scan in the drawing folder
    Dim exts(1) As String
    exts(0) = "pdf"
    exts(1) = "dwg"

    Dim i        As Integer
    Dim fileName As String
    Dim srcPath  As String
    Dim destPath As String
    Dim ts       As String

    For i = 0 To 1
        fileName = Dir(folder & baseName & "-Rev*." & exts(i))
        Do While fileName <> ""
            If LCase(Left(fileName, Len(currentRoot))) <> LCase(currentRoot) Then
                If Not fso.FolderExists(histFolder) Then fso.CreateFolder histFolder
                srcPath  = folder & fileName
                destPath = histFolder & fileName
                If fso.FileExists(destPath) Then
                    ts       = Format(Now, "YYYYMMDD_HHmmss")
                    destPath = histFolder & fso.GetBaseName(destPath) & "_" & ts & "." & fso.GetExtensionName(destPath)
                End If
                fso.MoveFile srcPath, destPath
            End If
            fileName = Dir()
        Loop
    Next i

    ' DXF sub-folder
    Dim dxfFolder As String
    dxfFolder = folder & "DXF\"
    If fso.FolderExists(dxfFolder) Then
        fileName = Dir(dxfFolder & baseName & "-Rev*.dxf")
        Do While fileName <> ""
            If LCase(Left(fileName, Len(currentRoot))) <> LCase(currentRoot) Then
                If Not fso.FolderExists(histFolder) Then fso.CreateFolder histFolder
                srcPath  = dxfFolder & fileName
                destPath = histFolder & fileName
                If fso.FileExists(destPath) Then
                    ts       = Format(Now, "YYYYMMDD_HHmmss")
                    destPath = histFolder & fso.GetBaseName(destPath) & "_" & ts & "." & fso.GetExtensionName(destPath)
                End If
                fso.MoveFile srcPath, destPath
            End If
            fileName = Dir()
        Loop
    End If

    Set fso = Nothing
End Sub

'==============================================================================
' ENSURE DXF FOLDER
'==============================================================================
Private Function EnsureDXFFolder(ByVal drawingFolder As String) As String
    Dim dxfPath As String
    dxfPath = drawingFolder & "DXF\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dxfPath) Then fso.CreateFolder dxfPath
    Set fso = Nothing
    EnsureDXFFolder = dxfPath
End Function

'==============================================================================
' EXPORT – PDF  (all sheets)
'==============================================================================
Private Function ExportToPDF(ByVal swDraw As SldWorks.DrawingDoc, _
                             ByVal outPath As String, _
                             ByRef errors As Long, _
                             ByRef warnings As Long) As Boolean
    Dim swApp      As SldWorks.SldWorks
    Dim swModel    As SldWorks.ModelDoc2
    Dim exportData As SldWorks.ExportPdfData

    Set swApp      = Application.SldWorks
    Set swModel    = swDraw
    Set exportData = swApp.GetExportFileData(swExportPdfData)

    exportData.ExportAs3D = False
    Dim sheetNames As Variant
    sheetNames = swDraw.GetSheetNames
    exportData.SetSheets swExportData_ExportAllSheets, sheetNames

    ExportToPDF = swModel.Extension.SaveAs(outPath, _
                                           swSaveAsCurrentVersion, _
                                           swSaveAsOptions_Silent, _
                                           exportData, errors, warnings)
End Function

'==============================================================================
' EXPORT – DWG
'==============================================================================
Private Function ExportToDWG(ByVal swDraw As SldWorks.DrawingDoc, _
                             ByVal outPath As String, _
                             ByRef errors As Long, _
                             ByRef warnings As Long) As Boolean
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw
    ExportToDWG = swModel.Extension.SaveAs(outPath, _
                                           swSaveAsCurrentVersion, _
                                           swSaveAsOptions_Silent, _
                                           Nothing, errors, warnings)
End Function

'==============================================================================
' EXPORT – DXF
'==============================================================================
Private Function ExportToDXF(ByVal swDraw As SldWorks.DrawingDoc, _
                             ByVal outPath As String, _
                             ByRef errors As Long, _
                             ByRef warnings As Long) As Boolean
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw
    ExportToDXF = swModel.Extension.SaveAs(outPath, _
                                           swSaveAsCurrentVersion, _
                                           swSaveAsOptions_Silent, _
                                           Nothing, errors, warnings)
End Function
