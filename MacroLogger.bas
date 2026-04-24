Attribute VB_Name = "MacroLogger"
'==============================================================================
' MacroLogger.bas  –  reusable dual-file logger for SolidWorks VBA macros
'
' HOW TO USE
' ----------
' 1. In the VBA IDE, choose File > Import File and select this .bas file.
' 2. Call WriteLog from anywhere in your macro:
'
'       MacroLogger.WriteLog _
'           logDir    := "Z:\DAG\SOLIDWORKS MACRO\My Macro\Log\", _
'           logName   := "MyMacro", _
'           colHeaders := Array("Job Number", "Result", "Notes"), _
'           colValues  := Array(jobNum, result, notes)
'
'    Or without named args (same thing):
'       MacroLogger.WriteLog "Z:\...\Log\", "MyMacro", _
'                            Array("Job Number", "Result"), _
'                            Array(jobNum, result)
'
' FILES PRODUCED
' --------------
'   Primary  : <logDir>\<logName>_Log.xlsx
'   Overflow : <logDir>\<logName>_Overflow.csv
'
'   The overflow CSV is written automatically whenever the .xlsx cannot be
'   opened (e.g. another user already has it open).
'
' SHEET LAYOUT
' ------------
'   Row 1 : "Total Runs" | =COUNTA(A4:A1048576)   ← bold summary line
'   Row 2 : (empty gap)
'   Row 3 : Date | Time | User | <your headers>    ← bold column headers
'   Row 4+: one row per WriteLog call
'
'   Date  = YYYY-MM-DD
'   Time  = HH:MM:SS
'   User  = Windows login name (Environ("USERNAME"))
'
' NOTES
' -----
' - colHeaders and colValues must be parallel arrays of the same length.
' - All values are written as-is via CStr(); format them before passing in.
' - The log directory is created automatically if it does not exist.
' - Uses late-bound COM (no Excel reference required in the VBA project).
'==============================================================================
Option Explicit

Public Sub WriteLog(ByVal logDir     As String, _
                    ByVal logName    As String, _
                    ByVal colHeaders As Variant, _
                    ByVal colValues  As Variant)

    ' Ensure trailing backslash
    If Right(logDir, 1) <> "\" Then logDir = logDir & "\"

    Dim logXlsx     As String : logXlsx     = logDir & logName & "_Log.xlsx"
    Dim logOverflow As String : logOverflow = logDir & logName & "_Overflow.csv"

    Const HEADER_ROW As Long = 3
    Const DATA_START As Long = 4

    Dim numCols As Long
    numCols = UBound(colHeaders) - LBound(colHeaders) + 1

    ' --- Ensure log directory exists ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logDir) Then fso.CreateFolder logDir

    ' --- Attempt primary Excel log ---
    Dim xlApp As Object
    Dim xlWB  As Object
    Dim xlWS  As Object

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then GoTo Overflow

    xlApp.Visible       = False
    xlApp.DisplayAlerts = False

    Dim lastRow As Long

    If fso.FileExists(logXlsx) Then
        ' Open existing file
        On Error Resume Next
        Set xlWB = xlApp.Workbooks.Open(logXlsx)
        If Err.Number <> 0 Then
            On Error GoTo 0
            xlApp.Quit : Set xlApp = Nothing
            GoTo Overflow
        End If
        On Error GoTo 0
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
        ' Create new workbook
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Sheets(1)
        xlWS.Name = logName & " Log"

        ' Summary row
        xlWS.Cells(1, 1).Value   = "Total Runs"
        xlWS.Cells(1, 2).Formula = "=COUNTA(A4:A1048576)"
        xlWS.Rows(1).Font.Bold   = True

        ' Header row: Date / Time / User are always first, then caller columns
        xlWS.Cells(HEADER_ROW, 1).Value = "Date"
        xlWS.Cells(HEADER_ROW, 2).Value = "Time"
        xlWS.Cells(HEADER_ROW, 3).Value = "User"
        Dim h As Long
        For h = 0 To numCols - 1
            xlWS.Cells(HEADER_ROW, 4 + h).Value = CStr(colHeaders(LBound(colHeaders) + h))
        Next h
        xlWS.Rows(HEADER_ROW).Font.Bold = True

        lastRow = DATA_START
    End If

    ' Write data row
    xlWS.Cells(lastRow, 1).Value = Format(Now, "YYYY-MM-DD")
    xlWS.Cells(lastRow, 2).Value = Format(Now, "HH:MM:SS")
    xlWS.Cells(lastRow, 3).Value = Environ("USERNAME")
    Dim v As Long
    For v = 0 To numCols - 1
        xlWS.Cells(lastRow, 4 + v).Value = colValues(LBound(colValues) + v)
    Next v

    ' Keep summary formula live – upgrades any file that previously had a hardcoded count
    xlWS.Cells(1, 2).Formula = "=COUNTA(A4:A1048576)"

    ' AutoFit all used columns
    Dim totalCols As Long : totalCols = 3 + numCols
    xlWS.Range(xlWS.Cells(1, 1), xlWS.Cells(lastRow, totalCols)).Columns.AutoFit

    If fso.FileExists(logXlsx) Then
        xlWB.Save
    Else
        xlWB.SaveAs logXlsx, 51    ' 51 = xlOpenXMLWorkbook (.xlsx)
    End If

    xlWB.Close False
    xlApp.Quit
    Set xlWS = Nothing : Set xlWB = Nothing : Set xlApp = Nothing
    Set fso = Nothing
    Exit Sub

    ' --- Overflow: .xlsx is locked, append to CSV instead ---
Overflow:
    Dim fn As Integer : fn = FreeFile

    If Not fso.FileExists(logOverflow) Then
        ' Write CSV header line
        Open logOverflow For Output As #fn
        Dim hdrLine As String : hdrLine = "Date,Time,User"
        Dim hi As Long
        For hi = LBound(colHeaders) To UBound(colHeaders)
            hdrLine = hdrLine & "," & CStr(colHeaders(hi))
        Next hi
        Print #fn, hdrLine
        Close #fn
    End If

    fn = FreeFile
    Open logOverflow For Append As #fn
    Dim dataLine As String
    dataLine = Format(Now, "YYYY-MM-DD") & "," & Format(Now, "HH:MM:SS") & "," & Environ("USERNAME")
    Dim vi As Long
    For vi = LBound(colValues) To UBound(colValues)
        dataLine = dataLine & "," & CStr(colValues(vi))
    Next vi
    Print #fn, dataLine
    Close #fn

    Set fso = Nothing
End Sub
