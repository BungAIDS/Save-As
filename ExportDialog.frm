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
Attribute VB_Name = "ExportDialog"
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
'   lblTitle        Label       "Save-As Export Utility"  (bold, large)
'   lblDrawing      Label       "Drawing:"
'   lblDrawingName  Label       (populated at runtime with drawing base name)
'   lblRevision     Label       "Revision Letter:"
'   txtRevision     TextBox     (user types A, B, C …)
'   fraFormats      Frame       "Export Formats"
'     chkPDF        CheckBox    "PDF"
'     chkDWG        CheckBox    "DWG"
'     chkDXF        CheckBox    "DXF"
'   lblPreview      Label       "Output file name preview:"
'   lblPreviewVal   Label       (populated at runtime)
'   btnOK           CommandButton  "Export"
'   btnCancel       CommandButton  "Cancel"
'==============================================================================
Option Explicit

' Public properties read by the caller
Public DrawingBaseName As String
Public RevisionLetter  As String
Public ExportPDF       As Boolean
Public ExportDWG       As Boolean
Public ExportDXF       As Boolean
Public Cancelled       As Boolean

'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Cancelled = False

    ' Pre-tick PDF as the most common choice
    chkPDF.Value = True
    chkDWG.Value = False
    chkDXF.Value = False

    ' Show drawing name in the label
    lblDrawingName.Caption = DrawingBaseName

    ' Clear preview
    UpdatePreview
End Sub

'------------------------------------------------------------------------------
' Live preview of the output filename as the user types the revision
'------------------------------------------------------------------------------
Private Sub txtRevision_Change()
    UpdatePreview
End Sub

Private Sub chkPDF_Click()  : UpdatePreview : End Sub
Private Sub chkDWG_Click()  : UpdatePreview : End Sub
Private Sub chkDXF_Click()  : UpdatePreview : End Sub

Private Sub UpdatePreview()
    Dim rev As String
    rev = UCase(Trim(txtRevision.Text))

    If rev = "" Then
        lblPreviewVal.Caption = "(enter a revision letter above)"
        Exit Sub
    End If

    Dim root As String
    root = DrawingBaseName & "-Rev" & rev

    Dim parts As String
    parts = ""

    If chkPDF.Value Then parts = parts & root & ".pdf" & vbCrLf
    If chkDWG.Value Then parts = parts & root & ".dwg" & vbCrLf
    If chkDXF.Value Then parts = parts & "DXF\" & root & ".dxf" & vbCrLf

    If parts = "" Then
        lblPreviewVal.Caption = "(select at least one format)"
    Else
        lblPreviewVal.Caption = parts
    End If
End Sub

'------------------------------------------------------------------------------
' OK / Export button
'------------------------------------------------------------------------------
Private Sub btnOK_Click()
    ' Validate revision letter - must be a single alpha character
    Dim rev As String
    rev = UCase(Trim(txtRevision.Text))

    If Len(rev) = 0 Then
        MsgBox "Please enter a revision letter (e.g. A, B, C).", _
               vbExclamation, "Save-As Export"
        txtRevision.SetFocus
        Exit Sub
    End If

    If Len(rev) > 2 Then
        MsgBox "Revision should be one or two characters (e.g. A or AA).", _
               vbExclamation, "Save-As Export"
        txtRevision.SetFocus
        Exit Sub
    End If

    ' Check that all chars are alphabetic
    Dim i As Integer
    For i = 1 To Len(rev)
        Dim c As String
        c = Mid(rev, i, 1)
        If (c < "A" Or c > "Z") Then
            MsgBox "Revision must contain only letters (A-Z).", _
                   vbExclamation, "Save-As Export"
            txtRevision.SetFocus
            Exit Sub
        End If
    Next i

    If Not chkPDF.Value And Not chkDWG.Value And Not chkDXF.Value Then
        MsgBox "Please select at least one export format.", _
               vbExclamation, "Save-As Export"
        Exit Sub
    End If

    ' Store values for caller
    RevisionLetter = rev
    ExportPDF      = chkPDF.Value
    ExportDWG      = chkDWG.Value
    ExportDXF      = chkDXF.Value
    Cancelled      = False

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
