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
    Dim extPos As Integer
    extPos = InStrRev(drawingBaseName, ".")
    If extPos > 0 Then drawingBaseName = Left(drawingBaseName, extPos - 1)

    If drawingBaseName = "" Then
        MsgBox "Could not determine drawing name from path:" & vbCrLf & drawingPath, _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

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
    Dim doSTP As Boolean
    doPDF = dlg.ExportPDF
    doDWG = dlg.ExportDWG
    doDXF = dlg.ExportDXF
    doSTP = dlg.ExportSTP

    Dim doEmail As Boolean
    doEmail = dlg.ExportEmail

    If Not doPDF And Not doDWG And Not doDXF And Not doSTP Then
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

    If doSTP Then
        Dim stepFolder As String
        stepFolder = EnsureSTEPFolder(acJobFolder)
        outPath = stepFolder & exportBase & ".step"
        If ClearToWrite(outPath) Then
            Dim swRefModel As SldWorks.ModelDoc2
            Set swRefModel = GetDrawingModel(swDraw)
            If swRefModel Is Nothing Then
                MsgBox "Could not find the referenced assembly or part for STEP export." & vbCrLf & _
                       "Make sure the model is open and loaded.", _
                       vbExclamation, "Save-As Export"
            Else
                ok = ExportToSTEP(swApp, swRefModel, outPath, errors, warnings)
                If ok Then
                    results = results & "  STEP: " & outPath & vbCrLf
                Else
                    MsgBox "STEP export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                           vbExclamation, "Save-As Export"
                End If
            End If
        End If
    End If

    If results <> "" Then
        MsgBox "Export complete!" & vbCrLf & vbCrLf & results, vbInformation, "Save-As Export"
        ' Open the AutoCAD job folder in Windows Explorer
        Shell "explorer.exe """ & acJobFolder & """", vbNormalFocus
        ' Log this run to the shared log
        LogExport jobNumber, drawingBaseName, swJobType, doPDF, doDWG, doDXF
        ' Draft transmittal e-mail if requested
        If doEmail Then DraftTransmittalEmail exportBase, revLetter
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

    ' --- STEP in the 3D STEP FILE sub-folder → 3D STEP FILE\History\ ---
    Dim stepFolder2    As String
    Dim stepHistFolder As String
    stepFolder2    = folder & "3D STEP FILE\"
    stepHistFolder = stepFolder2 & "HISTORY\"

    If fso.FolderExists(stepFolder2) Then
        fileName = Dir(stepFolder2 & baseNoRev & "*.step")
        Do While fileName <> ""
            If LCase(fso.GetBaseName(fileName)) <> LCase(currentBase) Then
                If Not fso.FolderExists(stepHistFolder) Then fso.CreateFolder stepHistFolder
                srcPath  = stepFolder2 & fileName
                destPath = stepHistFolder & fileName
                If fso.FileExists(destPath) Then
                    ts       = Format(Now, "YYYYMMDD_HHmmss")
                    destPath = stepHistFolder & fso.GetBaseName(fileName) & "_" & ts & ".step"
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
' Primary log: SaveAs_Log.xlsx  (Excel, full formatting)
' Overflow log: SaveAs_Overflow.csv  (used when xlsx is open/locked by a human)
'==============================================================================
Private Sub LogExport(ByVal jobNumber As String, _
                      ByVal drawingName As String, _
                      ByVal jobType As String, _
                      ByVal didPDF As Boolean, _
                      ByVal didDWG As Boolean, _
                      ByVal didDXF As Boolean)

    Const LOG_DIR      As String = "Z:\DAG\SOLIDWORKS MACRO\Save As\Log\"
    Const LOG_XLSX     As String = "Z:\DAG\SOLIDWORKS MACRO\Save As\Log\SaveAs_Log.xlsx"
    Const LOG_OVERFLOW As String = "Z:\DAG\SOLIDWORKS MACRO\Save As\Log\SaveAs_Overflow.csv"
    Const HEADER_ROW   As Long   = 3
    Const DATA_START   As Long   = 4

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(LOG_DIR) Then fso.CreateFolder LOG_DIR

    '--- Try primary Excel log first ---
    Dim xlApp As Object
    Dim xlWB  As Object
    Dim xlWS  As Object

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xlApp Is Nothing Then
        On Error GoTo 0
        GoTo Overflow
    End If
    On Error GoTo 0

    xlApp.Visible       = False
    xlApp.DisplayAlerts = False

    Dim lastRow As Long

    If fso.FileExists(LOG_XLSX) Then
        On Error Resume Next
        Set xlWB = xlApp.Workbooks.Open(LOG_XLSX)
        If Err.Number <> 0 Then
            ' File is locked – fall through to overflow
            On Error GoTo 0
            xlApp.Quit
            Set xlApp = Nothing
            GoTo Overflow
        End If
        On Error GoTo 0
        ' Some Excel versions open locked files as read-only instead of erroring
        If xlWB.ReadOnly Then
            xlWB.Close False
            xlApp.Quit
            Set xlWB = Nothing : Set xlApp = Nothing
            GoTo Overflow
        End If
        Set xlWS = xlWB.Sheets(1)
        lastRow = xlWS.Cells(xlWS.Rows.Count, 1).End(-4162).Row + 1
        If lastRow < DATA_START Then lastRow = DATA_START
    Else
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Sheets(1)
        xlWS.Name = "Export Log"

        xlWS.Cells(1, 1).Value = "Total Runs"
        xlWS.Cells(1, 2).Value = 0
        xlWS.Cells(1, 3).Value = "Time Saved"
        ' Formula auto-calculates time saved from B1 (total runs = total minutes)
        Dim tsf As String
        tsf = "=IF(B1=0,""0 minutes""," & _
              "IF(INT(B1/480)>0,INT(B1/480)&IF(INT(B1/480)=1,"" working day (8 hours each)"","" working days (8 hours each)"")&IF(INT(MOD(B1,480)/60)+MOD(B1,60)>0,"", "",""),"""")&" & _
              "IF(INT(MOD(B1,480)/60)>0,INT(MOD(B1,480)/60)&IF(INT(MOD(B1,480)/60)=1,"" hour"","" hours"")&IF(MOD(B1,60)>0,"", "",""),"""")&" & _
              "IF(MOD(B1,60)>0,MOD(B1,60)&IF(MOD(B1,60)=1,"" minute"","" minutes""),""""))"
        xlWS.Cells(1, 4).Formula = tsf
        xlWS.Rows(1).Font.Bold = True

        xlWS.Cells(HEADER_ROW, 1).Value = "Date"
        xlWS.Cells(HEADER_ROW, 2).Value = "Time"
        xlWS.Cells(HEADER_ROW, 3).Value = "User"
        xlWS.Cells(HEADER_ROW, 4).Value = "Job Number"
        xlWS.Cells(HEADER_ROW, 5).Value = "Drawing"
        xlWS.Cells(HEADER_ROW, 6).Value = "Job Type"
        xlWS.Cells(HEADER_ROW, 7).Value = "PDF"
        xlWS.Cells(HEADER_ROW, 8).Value = "DWG"
        xlWS.Cells(HEADER_ROW, 9).Value = "DXF"
        xlWS.Rows(HEADER_ROW).Font.Bold = True
        lastRow = DATA_START
    End If

    ' Write data row
    xlWS.Cells(lastRow, 1).Value = Format(Now, "YYYY-MM-DD")
    xlWS.Cells(lastRow, 2).Value = Format(Now, "HH:MM:SS")
    xlWS.Cells(lastRow, 3).Value = Environ("USERNAME")
    xlWS.Cells(lastRow, 4).Value = jobNumber
    xlWS.Cells(lastRow, 5).Value = drawingName
    xlWS.Cells(lastRow, 6).Value = jobType
    xlWS.Cells(lastRow, 7).Value = IIf(didPDF, "YES", "NO")
    xlWS.Cells(lastRow, 8).Value = IIf(didDWG, "YES", "NO")
    xlWS.Cells(lastRow, 9).Value = IIf(didDXF, "YES", "NO")

    ' Update total runs – D1 formula recalculates automatically
    Dim totalRuns As Long
    totalRuns = lastRow - DATA_START + 1
    xlWS.Cells(1, 2).Value = totalRuns
    xlWS.Columns("A:I").AutoFit

    If fso.FileExists(LOG_XLSX) Then
        xlWB.Save
    Else
        xlWB.SaveAs LOG_XLSX, 51
    End If

    xlWB.Close False
    xlApp.Quit
    Set xlWS = Nothing : Set xlWB = Nothing : Set xlApp = Nothing
    Set fso = Nothing
    Exit Sub

    '--- Overflow: xlsx was locked, append to CSV instead ---
Overflow:
    Dim fn As Integer
    fn = FreeFile

    If Not fso.FileExists(LOG_OVERFLOW) Then
        Open LOG_OVERFLOW For Output As #fn
        Print #fn, "Date,Time,User,Job Number,Drawing,Job Type,PDF,DWG,DXF"
        Close #fn
    End If

    fn = FreeFile
    Open LOG_OVERFLOW For Append As #fn
    Print #fn, Format(Now, "YYYY-MM-DD") & "," & _
               Format(Now, "HH:MM:SS") & "," & _
               Environ("USERNAME") & "," & _
               jobNumber & "," & drawingName & "," & jobType & "," & _
               IIf(didPDF, "YES", "NO") & "," & _
               IIf(didDWG, "YES", "NO") & "," & _
               IIf(didDXF, "YES", "NO")
    Close #fn

    Set fso = Nothing
End Sub

'==============================================================================
' DRAFT TRANSMITTAL E-MAIL
' Opens a new Outlook e-mail addressed to Debbie Decker.
' Subject = job number; body uses "order" or "revision" based on revLetter.
'==============================================================================
Private Sub DraftTransmittalEmail(ByVal exportBase As String, ByVal revLetter As String)
    If Trim(exportBase) = "" Then
        MsgBox "Cannot draft transmittal email: drawing name is empty.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    Dim orderOrRev As String
    orderOrRev = IIf(revLetter = "", "order", "revision")

    Dim signOff As String
    signOff = GetSignOffName()

    Dim body As String
    body = "Hi Debbie," & vbCrLf & vbCrLf & _
           "This " & orderOrRev & " is ready for transmittal and close." & vbCrLf & vbCrLf & _
           "Thanks,"
    If signOff <> "" Then body = body & vbCrLf & signOff

    On Error GoTo EmailErr

    ' CreateObject launches the registered COM server for Outlook (classic)
    Dim olApp  As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")

    Set olMail = olApp.CreateItem(0)   ' 0 = olMailItem

    ' Resolve recipient against Exchange address book
    Dim recip As Object
    Set recip = olMail.Recipients.Add("ddecker@chicagoblower.com")
    recip.Resolve

    olMail.Subject = exportBase
    olMail.Body    = body
    olMail.Display   ' Pops up the compose window ready to review and send

    Set olMail = Nothing
    Set olApp  = Nothing
    Exit Sub

EmailErr:
    MsgBox "Could not open Outlook to draft the transmittal e-mail." & vbCrLf & _
           "Please make sure Outlook (classic) is installed.", _
           vbExclamation, "Save-As Export – E-mail Error"
    Set olMail = Nothing
    Set olApp  = Nothing
End Sub

'--- Maps Windows username to first-name sign-off ---
Private Function GetSignOffName() As String
    Select Case LCase(Environ("USERNAME"))
        Case "dgroth":    GetSignOffName = "Danny"
        Case "llee":      GetSignOffName = "Latrell"
        Case "somar":     GetSignOffName = "Syed"
        Case "tsledz":    GetSignOffName = "Ted"
        Case "csandoval": GetSignOffName = "Carlos"
        Case "jbolda":    GetSignOffName = "Justin"
        Case Else:        GetSignOffName = ""
    End Select
End Function

'==============================================================================
' ENSURE 3D STEP FILE FOLDER
'==============================================================================
Private Function EnsureSTEPFolder(ByVal jobFolder As String) As String
    Dim stepPath As String
    stepPath = jobFolder & "3D STEP FILE\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(stepPath) Then fso.CreateFolder stepPath
    Set fso = Nothing
    EnsureSTEPFolder = stepPath
End Function

'==============================================================================
' GET REFERENCED MODEL FROM DRAWING
' Returns the first referenced assembly or part found in the drawing views.
'==============================================================================
Private Function GetDrawingModel(ByVal swDraw As SldWorks.DrawingDoc) As SldWorks.ModelDoc2
    Dim view    As SldWorks.View
    Dim refDoc  As SldWorks.ModelDoc2
    Set view = swDraw.GetFirstView()        ' First view is the sheet itself
    If Not view Is Nothing Then Set view = view.GetNextView()  ' Skip to first real view
    Do While Not view Is Nothing
        Set refDoc = view.ReferencedDocument
        If Not refDoc Is Nothing Then
            Set GetDrawingModel = refDoc
            Exit Function
        End If
        Set view = view.GetNextView()
    Loop
    Set GetDrawingModel = Nothing
End Function

'==============================================================================
' EXPORT – STEP AP203
' For parts: saves a temp .sldprt copy first (collapses feature tree so
' feature names are stripped), then exports that to STEP.
' For assemblies: macro recorder revealed the correct approach –
'   1. SetUserPreferenceIntegerValue to choose save-as-part mode
'   2. SaveAs3 on the active ModelDoc2 (NOT IAssemblyDoc.SaveAsPart)
' Using ExteriorFaces mode merges all geometry into one anonymous body,
' stripping component names and assembly structure for IP protection.
'==============================================================================
Private Function ExportToSTEP(ByVal swApp As SldWorks.SldWorks, _
                               ByVal swModel As SldWorks.ModelDoc2, _
                               ByVal outPath As String, _
                               ByRef errors As Long, _
                               ByRef warnings As Long) As Boolean

    ExportToSTEP = False

    Dim tempPath As String
    tempPath = Environ("TEMP") & "\SW_STEP_" & Format(Now, "YYYYMMDDHHmmss") & ".sldprt"

    Dim saveOk As Boolean

    If swModel.GetType = 1 Then  ' 1 = swDocPART

        ' Part: silent copy SaveAs collapses feature tree
        saveOk = swModel.Extension.SaveAs(tempPath, _
                                           swSaveAsCurrentVersion, _
                                           swSaveAsOptions_Silent Or swSaveAsOptions_Copy, _
                                           Nothing, errors, warnings)

    Else  ' 2 = swDocASSEMBLY

        ' Assembly: set the save-as-part preference to ExteriorFaces, then
        ' call SaveAs3 (IModelDoc2 method) – exactly what SolidWorks does
        ' internally when you do File > Save As > Part via the UI.
        ' ExteriorFaces (1) merges all geometry into one anonymous body and
        ' strips component names – AllComponents (2) preserves structure.

        ' Save original preference so we can restore it afterwards
        Dim origPref As Long
        origPref = swApp.GetUserPreferenceIntegerValue( _
                       swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions)

        swApp.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions, _
            swSaveAsmAsPartOptions_e.swSaveAsmAsPart_ExteriorFaces

        ' Activate the assembly – SaveAs3 operates on the active document
        Dim prevTitle As String
        Dim activErr  As Long
        prevTitle = swApp.ActiveDoc.GetTitle
        swApp.ActivateDoc2 swModel.GetTitle, False, activErr

        ' SaveAs3 returns Long: 0 = success (unlike Extension.SaveAs Boolean)
        Dim saveResult As Long
        saveResult = swApp.ActiveDoc.SaveAs3(tempPath, 0, 2)  ' 0=current ver, 2=copy
        saveOk = (saveResult = 0)

        ' Restore preference and re-activate the drawing
        swApp.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swSaveAssemblyAsPartOptions, origPref
        swApp.ActivateDoc2 prevTitle, False, activErr

    End If

    If saveOk Then
        Dim tempDoc As SldWorks.ModelDoc2
        Set tempDoc = swApp.OpenDoc6(tempPath, swDocPART, swOpenDocOptions_Silent, "", errors, warnings)

        If Not tempDoc Is Nothing Then
            ExportToSTEP = tempDoc.Extension.SaveAs(outPath, _
                                                     swSaveAsCurrentVersion, _
                                                     swSaveAsOptions_Silent, _
                                                     Nothing, errors, warnings)
            swApp.CloseDoc tempPath
        End If

        On Error Resume Next
        Kill tempPath
        On Error GoTo 0
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
