'==============================================================================
' SaveAs_Export.bas
' SolidWorks 2025 VBA Macro
'
' PURPOSE: Export an open .SLDDRW to PDF, DWG, and/or DXF with a revision
'          suffix.  Old revisions are automatically archived to a "History"
'          sub-folder.  DXF files are saved to a dedicated "DXF" sub-folder
'          (created automatically if absent).
'
' USAGE:   Open or activate the drawing you want to export, then run this
'          macro from Tools > Macros > Run.
'
' INSTALL: In the SolidWorks VBA editor (Tools > Macros > Edit) paste this
'          entire file into a module, or add it as a new module.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Entry point called by SolidWorks
'------------------------------------------------------------------------------
Sub main()

    Dim swApp       As SldWorks.SldWorks
    Dim swModel     As SldWorks.ModelDoc2
    Dim swDraw      As SldWorks.DrawingDoc

    Set swApp = Application.SldWorks

    ' Make sure a drawing is active
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

    ' Get the full path of the drawing, e.g. C:\Jobs\12345\12345.SLDDRW
    Dim drawingPath As String
    drawingPath = swModel.GetPathName

    If drawingPath = "" Then
        MsgBox "The drawing has never been saved.  Please save it first.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    ' Parse name and folder
    Dim drawingFolder   As String   ' e.g. C:\Jobs\12345\
    Dim drawingBaseName As String   ' e.g. 12345  (no extension)

    drawingFolder   = Left(drawingPath, InStrRev(drawingPath, "\"))
    drawingBaseName = Mid(drawingPath, Len(drawingFolder) + 1)
    drawingBaseName = Left(drawingBaseName, InStrRev(drawingBaseName, ".") - 1)

    ' Show the export dialog
    Dim dlg As New ExportDialog
    dlg.DrawingBaseName = drawingBaseName
    dlg.Show

    ' If the user cancelled, stop here
    If dlg.Cancelled Then Exit Sub

    ' Collect user choices
    Dim revLetter   As String
    Dim doPDF       As Boolean
    Dim doDWG       As Boolean
    Dim doDXF       As Boolean

    revLetter = UCase(Trim(dlg.RevisionLetter))
    doPDF     = dlg.ExportPDF
    doDWG     = dlg.ExportDWG
    doDXF     = dlg.ExportDXF

    If revLetter = "" Then
        MsgBox "No revision letter entered.  Export cancelled.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    If Not doPDF And Not doDWG And Not doDXF Then
        MsgBox "No export format selected.  Please check at least one box.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    ' Build output file name root  e.g. "12345-RevA"
    Dim exportRoot As String
    exportRoot = drawingBaseName & "-Rev" & revLetter

    '--------------------------------------------------------------------------
    ' Archive any existing revision files to the History folder
    '--------------------------------------------------------------------------
    ArchiveOldRevisions drawingFolder, drawingBaseName, exportRoot

    '--------------------------------------------------------------------------
    ' Export selected formats
    '--------------------------------------------------------------------------
    Dim errors   As Long
    Dim warnings As Long
    Dim outPath  As String
    Dim ok       As Boolean

    If doPDF Then
        outPath = drawingFolder & exportRoot & ".pdf"
        ok = ExportToPDF(swDraw, outPath, errors, warnings)
        If Not ok Then
            MsgBox "PDF export failed." & vbCrLf & _
                   "Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If doDWG Then
        outPath = drawingFolder & exportRoot & ".dwg"
        ok = ExportToDWG(swDraw, outPath, errors, warnings)
        If Not ok Then
            MsgBox "DWG export failed." & vbCrLf & _
                   "Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    If doDXF Then
        Dim dxfFolder As String
        dxfFolder = EnsureDXFFolder(drawingFolder)
        outPath = dxfFolder & exportRoot & ".dxf"
        ok = ExportToDXF(swDraw, outPath, errors, warnings)
        If Not ok Then
            MsgBox "DXF export failed." & vbCrLf & _
                   "Errors: " & errors & "  Warnings: " & warnings, _
                   vbExclamation, "Save-As Export"
        End If
    End If

    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "Output folder: " & drawingFolder & vbCrLf & _
           "File root: " & exportRoot, _
           vbInformation, "Save-As Export"

End Sub

'==============================================================================
' ARCHIVE - move all previous revision files to <drawingFolder>\History\
'           Matches any file whose name starts with <baseName>-Rev
'           but is NOT the current revision we are about to write.
'==============================================================================
Private Sub ArchiveOldRevisions(ByVal folder As String, _
                                ByVal baseName As String, _
                                ByVal currentRoot As String)

    Dim histFolder As String
    histFolder = folder & "History\"

    ' Collect candidate files with supported extensions
    Dim exts(2) As String
    exts(0) = "*.pdf"
    exts(1) = "*.dwg"
    exts(2) = "*.dxf"

    Dim fso     As Object
    Dim fsoFile As Object
    Dim pattern As String
    Dim i       As Integer
    Dim movedAny As Boolean
    movedAny = False

    Set fso = CreateObject("Scripting.FileSystemObject")

    For i = 0 To 2
        pattern = folder & baseName & "-Rev*." & Mid(exts(i), 3)
        Dim fileName As String
        fileName = Dir(pattern)

        Do While fileName <> ""
            ' Skip the file we are about to create
            If LCase(Left(fileName, Len(currentRoot))) <> LCase(currentRoot) Then
                ' This is an old revision - move it
                If Not fso.FolderExists(histFolder) Then
                    fso.CreateFolder histFolder
                End If

                Dim srcPath  As String
                Dim destPath As String
                srcPath  = folder & fileName
                destPath = histFolder & fileName

                ' If a file with the same name already exists in History,
                ' append a timestamp so nothing is overwritten
                If fso.FileExists(destPath) Then
                    Dim ts As String
                    ts = Format(Now, "YYYYMMDD_HHmmss")
                    Dim ext As String
                    ext = fso.GetExtensionName(destPath)
                    Dim base2 As String
                    base2 = fso.GetBaseName(destPath)
                    destPath = histFolder & base2 & "_" & ts & "." & ext
                End If

                fso.MoveFile srcPath, destPath
                movedAny = True
            End If
            fileName = Dir()
        Loop
    Next i

    ' Also check the DXF sub-folder if it exists
    Dim dxfFolder As String
    dxfFolder = folder & "DXF\"
    If fso.FolderExists(dxfFolder) Then
        fileName = Dir(dxfFolder & baseName & "-Rev*.dxf")
        Do While fileName <> ""
            If LCase(Left(fileName, Len(currentRoot))) <> LCase(currentRoot) Then
                If Not fso.FolderExists(histFolder) Then
                    fso.CreateFolder histFolder
                End If
                srcPath  = dxfFolder & fileName
                destPath = histFolder & fileName
                If fso.FileExists(destPath) Then
                    ts = Format(Now, "YYYYMMDD_HHmmss")
                    ext  = fso.GetExtensionName(destPath)
                    base2 = fso.GetBaseName(destPath)
                    destPath = histFolder & base2 & "_" & ts & "." & ext
                End If
                fso.MoveFile srcPath, destPath
                movedAny = True
            End If
            fileName = Dir()
        Loop
    End If

    Set fso = Nothing

End Sub

'==============================================================================
' ENSURE DXF FOLDER - returns the path (with trailing \) to the DXF folder,
'                     creating it if necessary.
'==============================================================================
Private Function EnsureDXFFolder(ByVal drawingFolder As String) As String

    Dim dxfPath As String
    dxfPath = drawingFolder & "DXF\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dxfPath) Then
        fso.CreateFolder dxfPath
    End If

    Set fso = Nothing
    EnsureDXFFolder = dxfPath

End Function

'==============================================================================
' EXPORT - PDF
'==============================================================================
Private Function ExportToPDF(ByVal swDraw As SldWorks.DrawingDoc, _
                             ByVal outPath As String, _
                             ByRef errors As Long, _
                             ByRef warnings As Long) As Boolean

    Dim swModel     As SldWorks.ModelDoc2
    Dim exportData  As SldWorks.ExportPdfData
    Dim swApp       As SldWorks.SldWorks

    Set swApp   = Application.SldWorks
    Set swModel = swDraw

    Set exportData = swApp.GetExportFileData(swExportPdfData)

    exportData.ExportAs3D = False

    ' Export all sheets
    Dim sheetNames As Variant
    sheetNames = swDraw.GetSheetNames
    exportData.SetSheets swExportData_ExportAllSheets, sheetNames

    ExportToPDF = swModel.Extension.SaveAs(outPath, _
                                           swSaveAsCurrentVersion, _
                                           swSaveAsOptions_Silent, _
                                           exportData, _
                                           errors, _
                                           warnings)
End Function

'==============================================================================
' EXPORT - DWG
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
                                           Nothing, _
                                           errors, _
                                           warnings)
End Function

'==============================================================================
' EXPORT - DXF
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
                                           Nothing, _
                                           errors, _
                                           warnings)
End Function
