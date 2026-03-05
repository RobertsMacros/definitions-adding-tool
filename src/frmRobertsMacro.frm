Option Explicit

Private Sub UserForm_Initialize()
    '=== Load saved preferences ===
    Dim defsCol As String, clauseCols As String, rowRanges As String
    LoadUserChoices defsCol, clauseCols, rowRanges
    If defsCol <> "" Then Me.txtDefinitionsCol.Text = defsCol
    If clauseCols <> "" Then Me.txtClauseCols.Text = clauseCols
    Me.txtRowRanges.Text = IIf(rowRanges <> "", rowRanges, "2:500")
    Me.btnRunRobertsMacro.Caption = "Run Macro"
    Me.btnSaveChoices.Caption = "Save Choices"
    Me.btnClose.Caption = "Close"
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2

End Sub
Private Sub btnSaveChoices_Click()
    SaveUserChoices _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text
    MsgBox "Choices saved. They will auto-load next time.", vbInformation
End Sub
Private Sub btnRunRobertsMacro_Click()
    RunDefinitionsTool _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub
