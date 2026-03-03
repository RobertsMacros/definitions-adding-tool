Attribute VB_Name = "modLaunchAndCore"
'=======================
' Module: modLaunchAndCore
' Entry points: opens the UI and validates / dispatches user inputs.
'=======================
Option Explicit

' Assigned to a macro shortcut (Alt+F8 -> Launch_RobertsMacro_UI).
Public Sub Launch_RobertsMacro_UI()
    frmRobertsMacro.Show
End Sub

' Called by the UserForm when the user clicks "Run RobertsMacro".
' Validates each input field before passing to the core engine.
Public Sub RunRobertsMacro_FromUI( _
    ByVal defsColText As String, _
    ByVal clauseColsText As String, _
    ByVal rowRangesText As String)

    Dim defsColNumber As Long
    Dim arrClauseCols() As Long
    Dim arrRows() As Long

    If Trim(defsColText) = "" Then
        MsgBox "Please enter the definitions column (e.g. AL).", vbExclamation
        Exit Sub
    End If

    defsColNumber = ColumnLetterToNumber(defsColText)
    If defsColNumber = 0 Then
        MsgBox "Definitions column entry is invalid: " & defsColText, vbExclamation
        Exit Sub
    End If

    arrClauseCols = ParseColumnList(clauseColsText)
    If UBound(arrClauseCols) = 0 And arrClauseCols(0) = 0 Then
        MsgBox "Clause columns entry is invalid or empty.", vbExclamation
        Exit Sub
    End If

    arrRows = ParseRowList(rowRangesText)
    If UBound(arrRows) = 0 And arrRows(0) = 0 Then
        MsgBox "Row ranges entry is invalid or empty.", vbExclamation
        Exit Sub
    End If

    InsertRelevantDefinitionsUnderClauses_Param arrRows, arrClauseCols, defsColNumber

End Sub
