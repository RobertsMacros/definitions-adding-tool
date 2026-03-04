Attribute VB_Name = "modLaunchAndCore"
'=======================
' Module: modLaunchAndCore
' Entry points: opens the UI and validates / dispatches user inputs.
'=======================
Option Explicit

' Assigned to a macro shortcut (Alt+F8 -> Launch_RobertsMacro_UI).
' On Mac, shows the InputBox runner instead of the UserForm.
Public Sub Launch_RobertsMacro_UI()
    #If Mac Then
        Launch_RobertsMacro_MacSafe
    #Else
        frmRobertsMacro.Show
    #End If
End Sub

' Mac-safe runner: prompts via InputBox using pipe-separated defaults.
Public Sub Launch_RobertsMacro_MacSafe()
    Dim defsCol As String
    Dim clauseCols As String
    Dim rowRanges As String

    LoadUserChoices defsCol, clauseCols, rowRanges

    Dim defaultVal As String
    defaultVal = defsCol & " | " & clauseCols & " | " & rowRanges

    Dim response As String
    response = InputBox("Enter: defsCol | clauseCols | rowRanges" & vbNewLine & _
                        "Example: AL | H:AK | 2:500", _
                        "Robert's Definitions Tool", defaultVal)

    If response = "" Then Exit Sub

    Dim parts() As String
    parts = Split(response, "|")
    If UBound(parts) < 2 Then
        MsgBox "Please enter all three values separated by | (pipe).", _
               vbExclamation, "Robert's Definitions Tool"
        Exit Sub
    End If

    RunDefinitionsTool Trim$(parts(0)), Trim$(parts(1)), Trim$(parts(2))
End Sub

' Single public entry point for running the tool from any caller
' (UserForm, MacSafe runner, or a future API/test harness).
' Normalises inputs, validates, persists, then executes the core engine.
Public Sub RunDefinitionsTool(ByVal defsCol As String, _
                              ByVal clauseCols As String, _
                              ByVal rowRanges As String)
    defsCol = NormaliseCol(defsCol)
    clauseCols = NormaliseCols(clauseCols)
    rowRanges = NormaliseRowRanges(rowRanges)

    Dim errMsg As String
    errMsg = ValidateInputs(defsCol, clauseCols, rowRanges)
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation, "Robert's Definitions Tool"
        Exit Sub
    End If

    SaveUserChoices defsCol, clauseCols, rowRanges

    Dim arrRows() As Long
    Dim arrCols() As Long
    arrRows = ParseRowList(rowRanges)
    arrCols = ParseColumnList(clauseCols)
    InsertRelevantDefinitionsUnderClauses_Param arrRows, arrCols, ColumnLetterToNumber(defsCol)
End Sub
