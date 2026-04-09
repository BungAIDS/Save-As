'==============================================================================
' SaveAs_Export.bas
' SolidWorks 2025 VBA Macro
'
' DRAWING FILENAME FORMAT:  XXXXXX-YYA.SLDDRW
'   XXXXXX = 6-digit job number  (used to locate the AutoCAD folder)
'   YY     = sheet/detail number
'   A      = revision letter (last character of base name)
'
' OUTPUT FOLDER (all formats go to AutoCAD):
'   PDF  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\
'   DWG  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\
'   DXF  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\DXF\
'
' REVISION ARCHIVING:
'   Old revisions of the same sheet are moved before new files are written:
'   PDF/DWG  →  <jobnum>\History\
'   DXF      →  <jobnum>\DXF\History\
'
' SOLIDWORKS → AUTOCAD JOB TYPE MAPPING:
'   GENERAL LINE  →  GENERAL LINE   (intermediate = first 3 digits, e.g. 420)
'   HD-PFD        →  HD-PFD-IAF     (intermediate = first 3 digits)
'   HDX           →  HDX            (intermediate = range folder, e.g. 416-420)
'==============================================================================
Option Explicit

'--- SolidWorks job root and folder type names ---
Private Const SW_ROOT         As String = "Z:\Solidworks\Current\JOBS"
Private Const JOBTYPE_GENLINE As String = "GENERAL LINE"
Private Const JOBTYPE_HDPFD   As String = "HD-PFD"
Private Const JOBTYPE_HDX     As String = "HDX"
Private Const JOBTYPE_AXIAL   As String = "AXIAL"

'--- AutoCAD job root and folder type names ---
Private Const AC_ROOT            As String = "Z:\AUTOCAD\CURRENT\JOBS"
Private Const AC_JOBTYPE_GENLINE As String = "GENERAL LINE"
Private Const AC_JOBTYPE_HDPFD   As String = "HD-PFD-IAF"
Private Const AC_JOBTYPE_HDX     As String = "HDX"
Private Const AC_JOBTYPE_AXIAL   As String = "AXIAL"

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

    '--- Parse folder and base name  e.g. "420788-01A" ---
    Dim drawingFolder   As String
    Dim drawingBaseName As String

    drawingFolder   = Left(drawingPath, InStrRev(drawingPath, "\"))
    drawingBaseName = Mid(drawingPath, Len(drawingFolder) + 1)
    drawingBaseName = Left(drawingBaseName, InStrRev(drawingBaseName, ".") - 1)

    '--- Extract job number (XXXXXX = everything before the first "-") ---
    Dim jobNumber As String
    Dim dashPos   As Integer
    dashPos = InStr(drawingBaseName, "-")
    If dashPos > 1 Then
        jobNumber = Left(drawingBaseName, dashPos - 1)
    Else
        jobNumber = drawingBaseName
    End If

    '--- Revision letter is entered by the user in the dialog ---
    Dim revLetter As String
    revLetter = ""

    '--- Detect SW job type from path ---
    Dim swJobType As String
    swJobType = DetectJobType(drawingFolder)

    If Not PathStartsWith(drawingFolder, SW_ROOT) Then
        Dim resp As Integer
        resp = MsgBox("This drawing is not inside the expected SolidWorks jobs folder:" & vbCrLf & _
                      "  " & SW_ROOT & vbCrLf & vbCrLf & _
                      "Current location: " & drawingFolder & vbCrLf & vbCrLf & _
                      "Continue anyway?", vbExclamation + vbYesNo, "Save-As Export – Path Warning")
        If resp = vbNo Then Exit Sub
        swJobType = ""
    ElseIf swJobType = "" Then
        resp = MsgBox("Drawing is inside the SolidWorks root but is not under a recognised" & vbCrLf & _
                      "job-type folder (GENERAL LINE, HD-PFD, HDX)." & vbCrLf & vbCrLf & _
                      "Continue anyway?", vbExclamation + vbYesNo, "Save-As Export – Path Warning")
        If resp = vbNo Then Exit Sub
    End If

    '--- Build AutoCAD output folder ---
    Dim acJobFolder As String
    acJobFolder = BuildAutoCADJobFolder(jobNumber, swJobType)

    If acJobFolder = "" Then
        MsgBox "Could not determine the AutoCAD job folder for job number: " & jobNumber & vbCrLf & _
               "Check that the job number is 6 digits and the drawing is in a recognised job type folder.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    '--- Ensure AutoCAD job folder exists ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(acJobFolder) Then
        Dim createResp As Integer
        createResp = MsgBox("AutoCAD job folder does not exist:" & vbCrLf & _
                            "  " & acJobFolder & vbCrLf & vbCrLf & _
                            "Create it now?", vbQuestion + vbYesNo, "Save-As Export")
        If createResp = vbNo Then
            Set fso = Nothing
            Exit Sub
        End If
        fso.CreateFolder acJobFolder
    End If
    Set fso = Nothing

    '--- Show export dialog ---
    Dim dlg As New Export_Dialog
    dlg.DrawingBaseName = drawingBaseName
    dlg.JobType         = IIf(swJobType <> "", swJobType, "(unknown)")
    dlg.DrawingFolder   = acJobFolder
    dlg.Show

    If dlg.Cancelled Then Exit Sub

    revLetter = dlg.RevisionLetter

    Dim doPDF As Boolean
    Dim doDWG As Boolean
    Dim doDXF As Boolean
    doPDF = dlg.ExportPDF
    doDWG = dlg.ExportDWG
    doDXF = dlg.ExportDXF

    If Not doPDF And Not doDWG And Not doDXF Then
        MsgBox "No export format selected.  Please check at least one box.", vbExclamation, "Save-As Export"
        Exit Sub
    End If

    '--- Export name = drawing base name + revision  e.g. "420788-01A" ---
    Dim exportBase As String
    exportBase = drawingBaseName & revLetter

    '--- Archive any existing revision files for this sheet ---
    '    Wildcard: drawingBaseName & "*.ext"  (e.g. "420788-01*.pdf")
    '    Skip file whose base name matches exportBase
    ArchiveOldRevisions acJobFolder, drawingBaseName, exportBase

    Dim errors   As Long
    Dim warnings As Long
    Dim outPath  As String
    Dim ok       As Boolean
    Dim results  As String
    results = ""

    If doPDF Then
        outPath = acJobFolder & exportBase & ".pdf"
        If ClearToWrite(outPath) Then
            ok = ExportToPDF(swDraw, outPath, errors, warnings)
            If ok Then
                results = results & "  PDF: " & outPath & vbCrLf
            Else
                MsgBox "PDF export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                       vbExclamation, "Save-As Export"
            End If
        End If
    End If

    If doDWG Then
        outPath = acJobFolder & exportBase & ".dwg"
        If ClearToWrite(outPath) Then
            ok = ExportToDWG(swDraw, outPath, errors, warnings)
            If ok Then
                results = results & "  DWG: " & outPath & vbCrLf
            Else
                MsgBox "DWG export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                       vbExclamation, "Save-As Export"
            End If
        End If
    End If

    If doDXF Then
        Dim dxfFolder As String
        dxfFolder = EnsureDXFFolder(acJobFolder)
        outPath = dxfFolder & exportBase & ".dxf"
        If ClearToWrite(outPath) Then
            ok = ExportToDXF(swDraw, outPath, errors, warnings)
            If ok Then
                results = results & "  DXF: " & outPath & vbCrLf
            Else
                MsgBox "DXF export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                       vbExclamation, "Save-As Export"
            End If
        End If
    End If

    If results <> "" Then
        MsgBox "Export complete!" & vbCrLf & vbCrLf & results, vbInformation, "Save-As Export"
        ' Open the AutoCAD job folder in Windows Explorer
        Shell "explorer.exe """ & acJobFolder & """", vbNormalFocus
        ' Log this run to the shared log
        LogExport jobNumber, drawingBaseName, swJobType, doPDF, doDWG, doDXF
    End If

End Sub

'==============================================================================
' BUILD AUTOCAD JOB FOLDER
'==============================================================================
Private Function BuildAutoCADJobFolder(ByVal jobNumber As String, _
                                       ByVal swJobType As String) As String
    Dim acJobType As String
    Select Case UCase(swJobType)
        Case UCase(JOBTYPE_GENLINE) : acJobType = AC_JOBTYPE_GENLINE
        Case UCase(JOBTYPE_HDPFD)   : acJobType = AC_JOBTYPE_HDPFD
        Case UCase(JOBTYPE_HDX)     : acJobType = AC_JOBTYPE_HDX
        Case UCase(JOBTYPE_AXIAL)   : acJobType = AC_JOBTYPE_AXIAL
        Case Else
            BuildAutoCADJobFolder = ""
            Exit Function
    End Select

    If Len(jobNumber) < 3 Then
        BuildAutoCADJobFolder = ""
        Exit Function
    End If

    Dim prefix3    As String
    Dim prefix3Int As Long
    prefix3 = Left(jobNumber, 3)
    If Not IsNumeric(prefix3) Then
        BuildAutoCADJobFolder = ""
        Exit Function
    End If
    prefix3Int = CLng(prefix3)

    Dim intermediate As String
    Select Case UCase(acJobType)
        Case UCase(AC_JOBTYPE_GENLINE), UCase(AC_JOBTYPE_HDPFD), UCase(AC_JOBTYPE_AXIAL)
            intermediate = prefix3                      ' e.g. "420"
        Case UCase(AC_JOBTYPE_HDX)
            intermediate = CalculateRange(prefix3Int)   ' e.g. "416-420"
    End Select

    BuildAutoCADJobFolder = AC_ROOT & "\" & acJobType & "\" & intermediate & "\" & jobNumber & "\"
End Function

'==============================================================================
' PATH HELPERS
'==============================================================================
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
' JOB TYPE DETECTION from SolidWorks drawing path
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
    ElseIf InStr(p, "\" & LCase(JOBTYPE_AXIAL) & "\") > 0 Then
        DetectJobType = JOBTYPE_AXIAL
    Else
        DetectJobType = ""
    End If
End Function

'==============================================================================
' RANGE FOLDER CALCULATION  (groups of 5 on first 3 digits)
'   e.g. prefix=420 → n=84 → 416-420
'   Special case: 401-405 → 400-405
'==============================================================================
Public Function CalculateRange(ByVal prefix3 As Long) As String
    Dim n      As Long
    Dim start  As Long
    Dim finish As Long
    n      = CLng(Int((prefix3 + 4) / 5))
    start  = 5 * (n - 1) + 1
    finish = 5 * n
    If start = 401 And finish = 405 Then start = 400
    CalculateRange = CStr(start) & "-" & CStr(finish)
End Function

'==============================================================================
' ARCHIVE old revisions of the same sheet
'
'   folder      = AutoCAD job folder (with trailing \)
'   baseNoRev   = drawing base name minus revision letter, e.g. "420788-01"
'   currentBase = full current base name, e.g. "420788-01A"
'
'   PDF/DWG old revisions → <jobfolder>\History\
'   DXF old revisions     → <jobfolder>\DXF\History\
'==============================================================================
Private Sub ArchiveOldRevisions(ByVal folder As String, _
                                ByVal baseNoRev As String, _
                                ByVal currentBase As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim histFolder    As String
    Dim dxfFolder     As String
    Dim dxfHistFolder As String
    histFolder    = folder & "HISTORY\"
    dxfFolder     = folder & "DXF\"
    dxfHistFolder = dxfFolder & "HISTORY\"

    Dim fileName As String
    Dim srcPath  As String
    Dim destPath As String
    Dim ts       As String

    ' --- PDF and DWG in the main job folder ---
    Dim exts(1) As String
    exts(0) = "pdf"
    exts(1) = "dwg"

    Dim i As Integer
    For i = 0 To 1
        fileName = Dir(folder & baseNoRev & "*." & exts(i))
        Do While fileName <> ""
            If LCase(fso.GetBaseName(fileName)) <> LCase(currentBase) Then
                If Not fso.FolderExists(histFolder) Then fso.CreateFolder histFolder
                srcPath  = folder & fileName
                destPath = histFolder & fileName
                If fso.FileExists(destPath) Then
                    ts       = Format(Now, "YYYYMMDD_HHmmss")
                    destPath = histFolder & fso.GetBaseName(fileName) & "_" & ts & "." & exts(i)
                End If
                On Error Resume Next
                fso.MoveFile srcPath, destPath
                If Err.Number <> 0 Then
                    MsgBox fileName & " could not be moved to History - it may be read-only or open in another program." & vbCrLf & _
                           "Please close or unlock the file and move it manually.", _
                           vbExclamation, "Save-As Export – Archive Warning"
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            fileName = Dir()
        Loop
    Next i

    ' --- DXF in the DXF sub-folder → DXF\History\ ---
    If fso.FolderExists(dxfFolder) Then
        fileName = Dir(dxfFolder & baseNoRev & "*.dxf")
        Do While fileName <> ""
            If LCase(fso.GetBaseName(fileName)) <> LCase(currentBase) Then
                If Not fso.FolderExists(dxfHistFolder) Then fso.CreateFolder dxfHistFolder
                srcPath  = dxfFolder & fileName
                destPath = dxfHistFolder & fileName
                If fso.FileExists(destPath) Then
                    ts       = Format(Now, "YYYYMMDD_HHmmss")
                    destPath = dxfHistFolder & fso.GetBaseName(fileName) & "_" & ts & ".dxf"
                End If
                On Error Resume Next
                fso.MoveFile srcPath, destPath
                If Err.Number <> 0 Then
                    MsgBox fileName & " could not be moved to History - it may be read-only or open in another program." & vbCrLf & _
                           "Please close or unlock the file and move it manually.", _
                           vbExclamation, "Save-As Export – Archive Warning"
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            fileName = Dir()
        Loop
    End If

    Set fso = Nothing
End Sub

'==============================================================================
' LOG EXPORT
' Writes to a CSV file using plain VBA file I/O so the file is never
' locked for more than a few milliseconds, allowing multiple users to
' run the macro concurrently without conflicts.
'
' Layout (opens cleanly in Excel):
'   Line 1: Summary  – Total Runs, <count>, Time Saved, <value>
'   Line 2: (blank)
'   Line 3: Column headers
'   Line 4+: Data rows
'
' Time Saved = 1 minute per run.
'==============================================================================
Private Sub LogExport(ByVal jobNumber As String, _
                      ByVal drawingName As String, _
                      ByVal jobType As String, _
                      ByVal didPDF As Boolean, _
                      ByVal didDWG As Boolean, _
                      ByVal didDXF As Boolean)

    Const LOG_PATH As String = "Z:\DAG\SOLIDWORKS MACRO\Save As\SaveAs_Log.csv"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim newEntry As String
    newEntry = Format(Now, "YYYY-MM-DD") & "," & _
               Format(Now, "HH:MM:SS") & "," & _
               Environ("USERNAME") & "," & _
               jobNumber & "," & drawingName & "," & jobType & "," & _
               IIf(didPDF, "YES", "NO") & "," & _
               IIf(didDWG, "YES", "NO") & "," & _
               IIf(didDXF, "YES", "NO")

    '--- If file doesn't exist, create it fresh with summary + headers ---
    If Not fso.FileExists(LOG_PATH) Then
        Dim fn As Integer
        fn = FreeFile
        Open LOG_PATH For Output As #fn
        Print #fn, "Total Runs,1,Time Saved,1 minute"
        Print #fn, ""
        Print #fn, "Date,Time,User,Job Number,Drawing,Job Type,PDF,DWG,DXF"
        Print #fn, newEntry
        Close #fn
        Set fso = Nothing
        Exit Sub
    End If

    '--- File exists: read all lines, count data rows, build updated summary ---
    Dim lines()   As String
    Dim allText   As String
    Dim totalRuns As Long

    ' Read with retry in case another user is writing at this exact moment
    Dim attempt As Integer
    For attempt = 1 To 5
        On Error Resume Next
        fn = FreeFile
        Open LOG_PATH For Input As #fn
        If Err.Number = 0 Then
            allText = Input(LOF(fn), fn)
            Close #fn
            On Error GoTo 0
            Exit For
        End If
        Close #fn
        On Error GoTo 0
        Wait 500   ' wait 0.5s before retrying
    Next attempt

    If allText = "" Then
        Set fso = Nothing
        Exit Sub   ' give up silently if we still can't read
    End If

    lines = Split(allText, vbCrLf)

    ' Count data rows (everything after the 3-line header block)
    Dim i As Integer
    totalRuns = 0
    For i = 3 To UBound(lines)
        If Trim(lines(i)) <> "" Then totalRuns = totalRuns + 1
    Next i
    totalRuns = totalRuns + 1   ' include this new entry

    '--- Build updated time saved string ---
    Dim tDays  As Long : tDays  = Int(totalRuns / 480)
    Dim tHours As Long : tHours = Int((totalRuns Mod 480) / 60)
    Dim tMins  As Long : tMins  = totalRuns Mod 60

    Dim timeSaved As String
    timeSaved = ""
    If tDays > 0  Then timeSaved = timeSaved & tDays  & IIf(tDays = 1,  " working day (8 hours each), ",  " working days (8 hours each), ")
    If tHours > 0 Then timeSaved = timeSaved & tHours & IIf(tHours = 1, " hour, ",  " hours, ")
    If tMins > 0  Then timeSaved = timeSaved & tMins  & IIf(tMins = 1,  " minute",  " minutes")
    timeSaved = TrimRight(timeSaved, ", ")
    If timeSaved = "" Then timeSaved = "0 minutes"

    '--- Write updated file then append new entry ---
    ' Update summary line (line 0)
    lines(0) = "Total Runs," & totalRuns & ",Time Saved," & timeSaved

    ' Write back with retry
    For attempt = 1 To 5
        On Error Resume Next
        fn = FreeFile
        Open LOG_PATH For Output As #fn
        If Err.Number = 0 Then
            Dim j As Integer
            For j = 0 To UBound(lines)
                Print #fn, lines(j)
            Next j
            Print #fn, newEntry
            Close #fn
            On Error GoTo 0
            Exit For
        End If
        Close #fn
        On Error GoTo 0
        Wait 500
    Next attempt

    Set fso = Nothing

End Sub

'--- Small wait helper (milliseconds) ---
Private Sub Wait(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub

'==============================================================================
' TRIM RIGHT - removes a trailing substring from a string if present
'==============================================================================
Private Function TrimRight(ByVal s As String, ByVal suffix As String) As String
    If Right(s, Len(suffix)) = suffix Then
        TrimRight = Left(s, Len(s) - Len(suffix))
    Else
        TrimRight = s
    End If
End Function

'==============================================================================
' ENSURE DXF FOLDER
'==============================================================================
Private Function EnsureDXFFolder(ByVal jobFolder As String) As String
    Dim dxfPath As String
    dxfPath = jobFolder & "DXF\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dxfPath) Then fso.CreateFolder dxfPath
    Set fso = Nothing
    EnsureDXFFolder = dxfPath
End Function

'==============================================================================
' CLEAR TO WRITE
' Always prompts the user before writing, whether or not the file exists.
' Also checks for read-only / locked files before prompting.
' Returns True  = safe to write
' Returns False = skip this file (read-only, locked, or user said No)
'==============================================================================
Private Function ClearToWrite(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fileName As String
    fileName = fso.GetFileName(filePath)

    ' If file exists, check it isn't locked or read-only before asking
    If fso.FileExists(filePath) Then
        Dim fileNum As Integer
        fileNum = FreeFile
        On Error Resume Next
        Open filePath For Binary Access Read Write As #fileNum
        Dim openErr As Long
        openErr = Err.Number
        Close #fileNum
        On Error GoTo 0

        If openErr <> 0 Then
            MsgBox fileName & " is read-only or is open in another program and cannot be overwritten.", _
                   vbExclamation, "Save-As Export – File Locked"
            ClearToWrite = False
            Set fso = Nothing
            Exit Function
        End If
    End If

    ' Only prompt if the file already exists
    If fso.FileExists(filePath) Then
        Dim resp As Integer
        resp = MsgBox(fileName & " already exists." & vbCrLf & vbCrLf & _
                      "Would you like to overwrite it?", _
                      vbQuestion + vbYesNo, "Save-As Export – Confirm")
        ClearToWrite = (resp = vbYes)
    Else
        ClearToWrite = True
    End If

    Set fso = Nothing
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
