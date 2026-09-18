Attribute VB_Name = "CloseTieOutCheck"
' ============================================================
' CloseTieOutCheck.bas
' Month-end control: the supporting schedule must tie exactly
' to the general ledger control total before the period closes.
'
' Setup: sheet named "TieOut" with
'   B2 = account name
'   B3 = schedule total
'   B4 = general ledger control total
' A log sheet named "ValidationLog" is created automatically.
' ============================================================

Option Explicit

Private Const TOLERANCE As Double = 0.005   ' half a cent

Public Sub RunTieOutCheck()

    Dim wsIn As Worksheet, wsLog As Worksheet
    Dim accountName As String
    Dim scheduleTotal As Double, glTotal As Double, variance As Double
    Dim status As String, nextRow As Long

    On Error GoTo ErrHandler

    Set wsIn = ThisWorkbook.Worksheets("TieOut")

    accountName = CStr(wsIn.Range("B2").Value)
    scheduleTotal = CDbl(wsIn.Range("B3").Value)
    glTotal = CDbl(wsIn.Range("B4").Value)

    variance = scheduleTotal - glTotal

    If Abs(variance) <= TOLERANCE Then
        status = "PASS"
    Else
        status = "FAIL"
    End If

    Set wsLog = GetOrCreateLog()
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(nextRow, 2).Value = Environ$("USERNAME")
    wsLog.Cells(nextRow, 3).Value = accountName
    wsLog.Cells(nextRow, 4).Value = scheduleTotal
    wsLog.Cells(nextRow, 5).Value = glTotal
    wsLog.Cells(nextRow, 6).Value = variance
    wsLog.Cells(nextRow, 7).Value = status

    wsLog.Range(wsLog.Cells(nextRow, 1), wsLog.Cells(nextRow, 7)).Font.Bold = (status = "FAIL")
    wsLog.Columns("A:G").AutoFit

    If status = "PASS" Then
        MsgBox accountName & " ties to the general ledger." & vbCrLf & _
               "Variance: " & Format(variance, "$#,##0.00") & vbCrLf & _
               "This account is cleared to close.", vbInformation, "Tie-out passed"
    Else
        MsgBox "CLOSE BLOCKED for " & accountName & "." & vbCrLf & vbCrLf & _
               "Schedule total: " & Format(scheduleTotal, "$#,##0.00") & vbCrLf & _
               "GL control total: " & Format(glTotal, "$#,##0.00") & vbCrLf & _
               "Variance: " & Format(variance, "$#,##0.00") & vbCrLf & vbCrLf & _
               "Investigate and correct before locking the period.", _
               vbCritical, "Tie-out failed"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Tie-out check could not run: " & Err.Description & vbCrLf & _
           "Confirm the TieOut sheet exists and B3 and B4 contain numbers.", _
           vbExclamation, "Control error"

End Sub

Private Function GetOrCreateLog() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ValidationLog")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "ValidationLog"
        ws.Range("A1:G1").Value = Array("Timestamp", "User", "Account", _
                                        "Schedule Total", "GL Control Total", _
                                        "Variance", "Status")
        ws.Range("A1:G1").Font.Bold = True
    End If

    Set GetOrCreateLog = ws

End Function
