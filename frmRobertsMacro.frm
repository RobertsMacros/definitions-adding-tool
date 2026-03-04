Option Explicit

Private Sub UserForm_Initialize()
    Dim defsCol As String
    Dim clauseCols As String
    Dim rowRanges As String

    LoadUserChoices defsCol, clauseCols, rowRanges

    If defsCol <> "" Then Me.txtDefinitionsCol.Text = defsCol
    If clauseCols <> "" Then Me.txtClauseCols.Text = clauseCols
    If rowRanges <> "" Then Me.txtRowRanges.Text = rowRanges
End Sub

Private Sub btnSaveChoices_Click()
    SaveUserChoices _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text

    MsgBox "Choices saved. They will auto-load next time.", vbInformation
End Sub

Private Sub btnRunRobertsMacro_Click()
    RunRobertsMacro_FromUI _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text
End Sub

Private Sub btnReportBug_Click()
    Dim mailTo As String
    Dim subj As String
    Dim body As String

    mailTo = "robert.stevens@freshfields.com"
    subj = "Robert's Definitions Tool - Bug Report"
    body = "Describe the issue here..."

    Application.FollowHyperlink _
        "mailto:" & mailTo & _
        "?subject=" & URLEncode(subj) & _
        "&body=" & URLEncode(body)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
