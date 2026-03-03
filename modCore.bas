Attribute VB_Name = "modCore"
'=======================
' Module: modCore
' Core engine: reads definitions, matches terms in clause cells,
' and writes the "Relevant Definitions:" block back to each cell.
'=======================
Option Explicit

'========================================================
' Public entry point
'========================================================
Public Sub InsertRelevantDefinitionsUnderClauses_Param( _
    ByRef arrRows() As Long, _
    ByRef arrClauseCols() As Long, _
    ByVal byColDefsSource As Long)

    Const LABEL_TEXT As String = "Relevant Definitions:"

    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim defSourceText As String
    Dim lineText As String
    Dim i As Long

    Dim rowDefs As Object          ' key phrase -> full definition text (per row)
    Dim cellUsed As Object         ' key phrase -> full definition text (per cell)

    Dim keyPhrase As String
    Dim arrRowKeys() As String     ' all keys found in the definitions column for this row
    Dim arrUsedKeys() As String    ' keys matched in the current clause cell
    Dim tmp As Variant
    Dim cellText As String
    Dim resultText As String
    Dim labelPos As Long

    Dim idxRow As Long
    Dim idxCol As Long

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    For idxRow = LBound(arrRows) To UBound(arrRows)
        r = arrRows(idxRow)

        defSourceText = ws.Cells(r, byColDefsSource).Value
        If Trim$(defSourceText) = "" Then GoTo NextRow

        '--------------------------------
        ' Build definitions dictionary for this row
        '--------------------------------
        Set rowDefs = CreateObject("Scripting.Dictionary")
        rowDefs.CompareMode = vbTextCompare

        ' Normalise line endings
        defSourceText = Replace(defSourceText, vbCrLf, vbLf)
        defSourceText = Replace(defSourceText, vbCr, vbLf)
        defSourceText = Trim$(defSourceText)

        ' Determine which separator was used between definitions
        Dim sep As String
        If InStr(defSourceText, "||") > 0 Then
            sep = "||"
        ElseIf InStr(defSourceText, "|") > 0 Then
            sep = "|"
        Else
            sep = vbLf & vbLf
            Do While InStr(defSourceText, vbLf & vbLf & vbLf) > 0
                defSourceText = Replace(defSourceText, vbLf & vbLf & vbLf, vbLf & vbLf)
            Loop
        End If

        Dim defItems() As String
        defItems = Split(defSourceText, sep)

        For i = LBound(defItems) To UBound(defItems)
            lineText = Trim$(defItems(i))
            If lineText <> "" Then
                AddDefinitionsFromLine rowDefs, lineText
            End If
        Next i

        If rowDefs.Count = 0 Then GoTo NextRow

        ' Cache the keys so we don't re-enumerate the dictionary inside the column loop
        ReDim arrRowKeys(0 To rowDefs.Count - 1)
        i = 0
        For Each tmp In rowDefs.Keys
            arrRowKeys(i) = CStr(tmp)
            i = i + 1
        Next tmp

        '--------------------------------
        ' Process each clause column for this row
        '--------------------------------
        For idxCol = LBound(arrClauseCols) To UBound(arrClauseCols)
            c = arrClauseCols(idxCol)
            cellText = ws.Cells(r, c).Value

            If Trim$(cellText) <> "" Then

                ' Strip any previously inserted definitions block before matching,
                ' so we don't accidentally match terms inside old definitions.
                Dim cellTextForMatch As String
                labelPos = InStr(1, cellText, LABEL_TEXT, vbTextCompare)
                If labelPos > 0 Then
                    cellTextForMatch = Left$(cellText, labelPos - 1)
                Else
                    cellTextForMatch = cellText
                End If

                ' Flatten whitespace so phrase matching isn't broken by line breaks
                cellTextForMatch = Replace(cellTextForMatch, vbCrLf, " ")
                cellTextForMatch = Replace(cellTextForMatch, vbCr, " ")
                cellTextForMatch = Replace(cellTextForMatch, vbLf, " ")
                cellTextForMatch = Replace(cellTextForMatch, Chr(160), " ")

                Set cellUsed = CreateObject("Scripting.Dictionary")
                cellUsed.CompareMode = vbTextCompare

                For i = LBound(arrRowKeys) To UBound(arrRowKeys)
                    keyPhrase = arrRowKeys(i)
                    If PhraseInText(cellTextForMatch, keyPhrase) Then
                        If Not cellUsed.Exists(keyPhrase) Then
                            cellUsed.Add keyPhrase, rowDefs(keyPhrase)
                        End If
                    End If
                Next i

                If cellUsed.Count > 0 Then

                    ' Sort matched keys alphabetically (case-insensitive bubble sort)
                    ReDim arrUsedKeys(0 To cellUsed.Count - 1)
                    i = 0
                    For Each tmp In cellUsed.Keys
                        arrUsedKeys(i) = CStr(tmp)
                        i = i + 1
                    Next tmp

                    Dim j As Long, k As Long
                    For j = LBound(arrUsedKeys) To UBound(arrUsedKeys) - 1
                        For k = j + 1 To UBound(arrUsedKeys)
                            If UCase$(arrUsedKeys(j)) > UCase$(arrUsedKeys(k)) Then
                                tmp = arrUsedKeys(j)
                                arrUsedKeys(j) = arrUsedKeys(k)
                                arrUsedKeys(k) = tmp
                            End If
                        Next k
                    Next j

                    ' Build the definitions block
                    Dim defBlock As String
                    defBlock = LABEL_TEXT
                    For i = LBound(arrUsedKeys) To UBound(arrUsedKeys)
                        defBlock = defBlock & vbLf & vbLf & cellUsed(arrUsedKeys(i))
                    Next i

                    ' Replace or append the definitions block
                    labelPos = InStr(1, cellText, LABEL_TEXT, vbTextCompare)
                    If labelPos > 0 Then
                        resultText = Left$(cellText, labelPos - 1)
                        resultText = RTrim$(resultText) & vbLf & vbLf & defBlock
                    Else
                        resultText = cellText & vbLf & vbLf & defBlock
                    End If

                    ws.Cells(r, c).Value = resultText
                    ApplyFormattingToCell ws.Cells(r, c), LABEL_TEXT
                End If
            End If
        Next idxCol

NextRow:
    Next idxRow

    Application.ScreenUpdating = True
    MsgBox "Relevant definitions inserted/refreshed.", vbInformation
End Sub

'========================================================
' Private helpers
'========================================================

' Extract the key phrase and full text from one definition line and add to dict.
Private Sub AddDefinitionsFromLine(ByVal rowDefs As Object, ByVal lineText As String)
    Dim defText As String
    Dim keyPhrase As String

    defText = Trim$(lineText)
    If defText = "" Then Exit Sub

    keyPhrase = ExtractKeyPhrase_W(defText)
    If keyPhrase <> "" Then
        If Not rowDefs.Exists(keyPhrase) Then
            rowDefs.Add keyPhrase, defText
        End If
    End If
End Sub

' Return the defined term from a definition line of the form "<Term> means ...".
Private Function ExtractKeyPhrase_W(ByVal lineText As String) As String
    Dim posMeans As Long
    Dim candidate As String

    posMeans = InStr(1, LCase$(lineText), " means")
    If posMeans = 0 Then Exit Function

    candidate = Left$(lineText, posMeans - 1)
    ExtractKeyPhrase_W = StripPunctuation(Trim$(candidate))
End Function

' Remove leading and trailing non-alphanumeric characters from a string.
Private Function StripPunctuation(ByVal txt As String) As String
    Dim ch As String

    Do While Len(txt) > 0
        ch = Left$(txt, 1)
        If ch Like "[A-Za-z0-9]" Then Exit Do
        txt = Mid$(txt, 2)
    Loop

    Do While Len(txt) > 0
        ch = Right$(txt, 1)
        If ch Like "[A-Za-z0-9]" Then Exit Do
        txt = Left$(txt, Len(txt) - 1)
    Loop

    StripPunctuation = txt
End Function

' Return True when keyPhrase appears as a whole word (or with a plural-s) in txt.
'
' BUG FIX: the original code called Mid$(txt, pos - 1, 1) without guarding
' against pos = 1, which passes start=0 to Mid$ and raises runtime error 5
' ("Invalid procedure call or argument").  The fix reads the character before
' the match only when pos > 1, and treats the start-of-string as a word boundary.
Private Function PhraseInText(ByVal txt As String, ByVal keyPhrase As String) As Boolean
    Dim pos As Long
    Dim keyLen As Long
    Dim charBefore As String

    PhraseInText = False
    If txt = "" Or keyPhrase = "" Then Exit Function

    keyLen = Len(keyPhrase)
    pos = InStr(1, txt, keyPhrase, vbTextCompare)

    Do While pos > 0
        ' Safe read of the character immediately before the match.
        If pos > 1 Then
            charBefore = Mid$(txt, pos - 1, 1)
        Else
            charBefore = "" ' start of string counts as a word boundary
        End If

        If IsWordBoundary(charBefore, True) _
        And IsWordBoundary(Mid$(txt, pos + keyLen, 1), False) Then
            PhraseInText = True
            Exit Function
        End If
        pos = InStr(pos + 1, txt, keyPhrase, vbTextCompare)
    Loop
End Function

' Return True when ch is an acceptable word boundary character.
' isBefore=False permits a trailing "s" / "S" (handles simple plurals).
Private Function IsWordBoundary(ByVal ch As String, Optional ByVal isBefore As Boolean = False) As Boolean
    If ch = "" Then
        IsWordBoundary = True
    ElseIf ch = """" Then
        IsWordBoundary = True
    ElseIf Not isBefore And (ch = "s" Or ch = "S") Then
        IsWordBoundary = True
    ElseIf ch Like "[A-Za-z0-9]" Then
        IsWordBoundary = False
    Else
        IsWordBoundary = True
    End If
End Function

' Bold the "Relevant Definitions:" label text inside a cell.
Private Sub ApplyFormattingToCell(ByVal targetCell As Range, ByVal labelText As String)
    Dim fullText As String
    Dim labelPos As Long

    fullText = CStr(targetCell.Value)
    labelPos = InStr(1, fullText, labelText, vbTextCompare)
    If labelPos = 0 Then Exit Sub

    targetCell.Characters(labelPos, Len(labelText)).Font.Bold = True
End Sub
