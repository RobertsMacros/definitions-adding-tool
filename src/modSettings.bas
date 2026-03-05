Attribute VB_Name = "modSettings"
'=======================
' Module: modSettings
' Persists and retrieves user choices in a hidden worksheet so they
' survive workbook save/close cycles.
'=======================
Option Explicit

Private Const SHEET_NAME As String = "_RobertsMacroSettings"

' Save the three user inputs to the hidden settings sheet.
' Creates the sheet (very hidden) if it does not already exist.
Public Sub SaveUserChoices(ByVal defsCol As String, _
                           ByVal clauseCols As String, _
                           ByVal rowRanges As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_NAME
        ws.Visible = xlSheetVeryHidden
    End If

    ws.Range("A1").Value = "DefinitionsColumn"
    ws.Cells(1, 2).NumberFormat = "@" ' FIX: force text so e.g. "AL" is never parsed as a date/time
    ws.Cells(1, 2).Value = defsCol
    ws.Range("A2").Value = "ClauseColumns"
    ws.Cells(2, 2).NumberFormat = "@" ' FIX: force text so e.g. "H:AK" is never parsed as a date/time
    ws.Cells(2, 2).Value = clauseCols
    ws.Range("A3").Value = "RowRanges"
    ws.Cells(3, 2).NumberFormat = "@" ' FIX: force text so e.g. "5:20" is never stored as a time fraction
    ws.Cells(3, 2).Value = rowRanges

End Sub

' Load previously saved user inputs from the hidden settings sheet.
' Returns empty strings when the sheet does not exist yet.
Public Sub LoadUserChoices(ByRef defsCol As String, _
                           ByRef clauseCols As String, _
                           ByRef rowRanges As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        defsCol = ""
        clauseCols = ""
        rowRanges = ""
        Exit Sub
    End If

    defsCol    = ReadTextCell(ws.Cells(1, 2))
    clauseCols = ReadTextCell(ws.Cells(2, 2))
    rowRanges  = ReadTextCell(ws.Cells(3, 2))

End Sub

' Returns the cell's value as text, or "" if numeric (e.g. a time/date
' that Excel auto-parsed from a prior save without the "@" number format).
Private Function ReadTextCell(ByVal cell As Range) As String
    If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
        ReadTextCell = ""
    Else
        ReadTextCell = CStr(cell.Value)
    End If
End Function
