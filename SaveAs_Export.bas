'==============================================================================
' SaveAs_Export.bas
' SolidWorks 2025 VBA Macro
'
' PURPOSE: Export an open .SLDDRW (format XXXXXX-YY) to PDF, DWG, and/or DXF
'          with a revision suffix, saved to the corresponding AutoCAD job folder.
'
' DRAWING FILENAME FORMAT:  XXXXXX-YY.SLDDRW
'   XXXXXX = 6-digit job number  (used to locate the AutoCAD folder)
'   YY     = sheet/detail number
'
' OUTPUT FOLDER (all formats go to AutoCAD):
'   PDF  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\
'   DWG  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\
'   DXF  →  Z:\AUTOCAD\CURRENT\JOBS\<type>\<intermediate>\<jobnum>\DXF\
'
' SOLIDWORKS → AUTOCAD JOB TYPE MAPPING:
'   GENERAL LINE  →  GENERAL LINE   (intermediate = first 3 digits, e.g. 420)
'   HD-PFD        →  HD-PFD-IAF     (intermediate = first 3 digits)
'   HDX           →  HDX            (intermediate = range folder, e.g. 416-420)
'
' REVISION ARCHIVING:
'   Existing revision files in the AutoCAD job folder (and DXF sub-folder)
'   are moved to a History\ sub-folder before new files are written.
'==============================================================================
Option Explicit

'--- SolidWorks job root and folder type names ---
Private Const SW_ROOT          As String = "Z:\Solidworks\Current\JOBS"
Private Const JOBTYPE_GENLINE  As String = "GENERAL LINE"
Private Const JOBTYPE_HDPFD    As String = "HD-PFD"
Private Const JOBTYPE_HDX      As String = "HDX"

'--- AutoCAD job root and folder type names ---
Private Const AC_ROOT          As String = "Z:\AUTOCAD\CURRENT\JOBS"
Private Const AC_JOBTYPE_GENLINE As String = "GENERAL LINE"
Private Const AC_JOBTYPE_HDPFD   As String = "HD-PFD-IAF"
Private Const AC_JOBTYPE_HDX     As String = "HDX"

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

    '--- Parse folder and full base name (e.g. "420788-01") ---
    Dim drawingFolder   As String
    Dim drawingBaseName As String

    drawingFolder   = Left(drawingPath, InStrRev(drawingPath, "\"))
    drawingBaseName = Mid(drawingPath, Len(drawingFolder) + 1)
    drawingBaseName = Left(drawingBaseName, InStrRev(drawingBaseName, ".") - 1)

    '--- Extract job number (part before the first "-") ---
    Dim jobNumber As String
    Dim dashPos   As Integer
    dashPos = InStr(drawingBaseName, "-")
    If dashPos > 1 Then
        jobNumber = Left(drawingBaseName, dashPos - 1)
    Else
        ' No dash found – treat entire base name as the job number
        jobNumber = drawingBaseName
    End If

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

    '--- Build the AutoCAD output folder from job number + job type ---
    Dim acJobFolder As String
    acJobFolder = BuildAutoCADJobFolder(jobNumber, swJobType)

    If acJobFolder = "" Then
        MsgBox "Could not determine the AutoCAD job folder for job number: " & jobNumber & vbCrLf & _
               "Check that the job number is 6 digits and the drawing is in a recognised job type folder.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    '--- Ensure the AutoCAD job folder exists ---
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

    '--- Build export root name  e.g. "420788-01-RevA" ---
    Dim exportRoot As String
    exportRoot = drawingBaseName & "-Rev" & revLetter

    '--- Archive previous revisions in the AutoCAD folder ---
    ArchiveOldRevisions acJobFolder, drawingBaseName, exportRoot

    '--- Export ---
    Dim errors   As Long
    Dim warnings As Long
    Dim outPath  As String
    Dim ok       As Boolean
    Dim results  As String
    results = ""

    If doPDF Then
        outPath = acJobFolder & exportRoot & ".pdf"
        ok = ExportToPDF(swDraw, outPath, errors, warnings)
        If ok Then
            results = results & "  PDF: " & outPath & vbCrLf
        Else
            MsgBox "PDF export failed.  Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If doDWG Then
        outPath = acJobFolder & exportRoot & ".dwg"
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
        dxfFolder = EnsureDXFFolder(acJobFolder)
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
' BUILD AUTOCAD JOB FOLDER
' Maps SW job type to AutoCAD job type, calculates the intermediate folder,
' and returns the full AutoCAD job folder path (with trailing backslash).
' Returns "" on error.
'==============================================================================
Private Function BuildAutoCADJobFolder(ByVal jobNumber As String, _
                                       ByVal swJobType As String) As String
    ' Map SW job type → AutoCAD job type
    Dim acJobType As String
    Select Case UCase(swJobType)
        Case UCase(JOBTYPE_GENLINE) : acJobType = AC_JOBTYPE_GENLINE
        Case UCase(JOBTYPE_HDPFD)   : acJobType = AC_JOBTYPE_HDPFD
        Case UCase(JOBTYPE_HDX)     : acJobType = AC_JOBTYPE_HDX
        Case Else
            BuildAutoCADJobFolder = ""
            Exit Function
    End Select

    ' Build intermediate folder
    Dim intermediate As String
    Dim prefix3      As String
    Dim prefix3Int   As Long

    If Len(jobNumber) < 3 Then
        BuildAutoCADJobFolder = ""
        Exit Function
    End If

    prefix3 = Left(jobNumber, 3)
    If Not IsNumeric(prefix3) Then
        BuildAutoCADJobFolder = ""
        Exit Function
    End If
    prefix3Int = CLng(prefix3)

    Select Case UCase(acJobType)
        Case UCase(AC_JOBTYPE_GENLINE), UCase(AC_JOBTYPE_HDPFD)
            ' AutoCAD GENERAL LINE and HD-PFD-IAF use first 3 digits as intermediate
            intermediate = prefix3

        Case UCase(AC_JOBTYPE_HDX)
            ' AutoCAD HDX uses range folder
            intermediate = CalculateRange(prefix3Int)
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
' ARCHIVE old revision files → <acJobFolder>\History\
' Scans the AutoCAD job folder and DXF sub-folder for any file matching
' <baseName>-Rev*.ext that is NOT the current revision, and moves them.
'==============================================================================
Private Sub ArchiveOldRevisions(ByVal folder As String, _
                                ByVal baseName As String, _
                                ByVal currentRoot As String)

    Dim histFolder As String
    histFolder = folder & "History\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim exts(1)  As String
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
' ENSURE DXF FOLDER  (creates DXF\ inside the AutoCAD job folder if absent)
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
