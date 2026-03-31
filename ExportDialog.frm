VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportDialog
   Caption         =   "Save-As Export"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "ExportDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Export_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' ExportDialog.frm
' UserForm for the SaveAs_Export macro.
'
' Controls (add these in the VBA UserForm designer):
'
'   lblTitle        Label          "Save-As Export Utility"  (bold, size 12)
'   lblJobType      Label          "Job Type:"
'   lblJobTypeVal   Label          (populated at runtime – e.g. "GENERAL LINE")
'   lblFolder       Label          "Save folder:"
'   lblFolderVal    Label          (populated at runtime – full path, WordWrap=True)
'   lblDrawing      Label          "Drawing / Job #:"
'   lblDrawingName  Label          (populated at runtime with drawing base name)
'   lblRevision     Label          "Revision Letter:"
'   txtRevision     TextBox        (user types A, B, C …  MaxLength=2)
'   fraFormats      Frame          "Export Formats"
'     chkPDF        CheckBox       "PDF (.pdf)"
'     chkDWG        CheckBox       "AutoCAD DWG (.dwg)"
'     chkDXF        CheckBox       "DXF (.dxf)  → saved in DXF\ sub-folder"
'   lblPreview      Label          "Output file name preview:"
'   lblPreviewVal   Label          (populated at runtime, WordWrap=True)
'   btnOK           CommandButton  "Export"   Default=True
'   btnCancel       CommandButton  "Cancel"   Cancel=True
'==============================================================================
Option Explicit

' Public properties – set by caller before Show, read back after Hide
Public DrawingBaseName As String
Public JobType         As String   ' e.g. "GENERAL LINE", "HD-PFD", "HDX"
Public DrawingFolder   As String   ' AutoCAD output folder (full path, trailing \)
Public RevisionLetter  As String   ' set by caller (read from filename); read-only in dialog
Public ExportPDF       As Boolean
Public ExportDWG       As Boolean
Public ExportDXF       As Boolean
Public Cancelled       As Boolean

'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Cancelled = False

    chkPDF.Value = True
    chkDWG.Value = False
    chkDXF.Value = False

    ' Populate read-only info labels
    lblDrawingName.Caption = DrawingBaseName
    lblJobTypeVal.Caption  = IIf(JobType <> "", JobType, "(not detected)")
    lblFolderVal.Caption   = IIf(DrawingFolder <> "", DrawingFolder, "(unknown)")

    ' Revision is read from the filename – show it but don't let user edit it
    ' (txtRevision is renamed to lblRevisionVal in the designer: change it to
    '  a Label with the same position, or set Enabled=False / Locked=True)
    txtRevision.Text    = RevisionLetter
    txtRevision.Enabled = False
    txtRevision.Locked  = True

    UpdatePreview
End Sub

'------------------------------------------------------------------------------
Private Sub chkPDF_Click() : UpdatePreview : End Sub
Private Sub chkDWG_Click() : UpdatePreview : End Sub
Private Sub chkDXF_Click() : UpdatePreview : End Sub

Private Sub UpdatePreview()
    Dim parts As String
    parts = ""

    If chkPDF.Value Then parts = parts & DrawingBaseName & ".pdf" & vbCrLf
    If chkDWG.Value Then parts = parts & DrawingBaseName & ".dwg" & vbCrLf
    If chkDXF.Value Then parts = parts & "DXF\" & DrawingBaseName & ".dxf" & vbCrLf

    If parts = "" Then
        lblPreviewVal.Caption = "(select at least one format)"
    Else
        lblPreviewVal.Caption = parts
    End If
End Sub

'------------------------------------------------------------------------------
' Export button
'------------------------------------------------------------------------------
Private Sub btnOK_Click()
    If Not chkPDF.Value And Not chkDWG.Value And Not chkDXF.Value Then
        MsgBox "Please select at least one export format.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    ExportPDF = chkPDF.Value
    ExportDWG = chkDWG.Value
    ExportDXF = chkDXF.Value
    Cancelled = False

    Me.Hide
End Sub

'------------------------------------------------------------------------------
' Cancel button
'------------------------------------------------------------------------------
Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

'------------------------------------------------------------------------------
' Close button (X) on the form header also counts as cancel
'------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Cancelled = True
        Me.Hide
    End If
End Sub
