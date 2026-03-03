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
    ws.Range("B1").Value = defsCol
    ws.Range("A2").Value = "ClauseColumns"
    ws.Range("B2").Value = clauseCols
    ws.Range("A3").Value = "RowRanges"
    ws.Range("B3").Value = rowRanges

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

    defsCol = CStr(ws.Range("B1").Value)
    clauseCols = CStr(ws.Range("B2").Value)
    rowRanges = CStr(ws.Range("B3").Value)

End Sub
