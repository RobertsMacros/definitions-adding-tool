Attribute VB_Name = "modHelpers"
'=======================
' Module: modHelpers
' Shared utility functions used across modules.
'=======================
Option Explicit

' Convert a column letter (or letters) to the corresponding column number.
' Examples: "A" -> 1, "Z" -> 26, "AA" -> 27, "AL" -> 38
Public Function ColumnLetterToNumber(ByVal colRef As String) As Long
    Dim s As String
    Dim i As Long
    Dim result As Long

    s = UCase$(Trim$(colRef))
    result = 0

    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "[A-Z]" Then
            result = result * 26 + (Asc(Mid$(s, i, 1)) - Asc("A") + 1)
        End If
    Next i

    ColumnLetterToNumber = result
End Function

' Normalise a raw textbox string so that ParseRowList / ParseColumnList
' can split it reliably on commas (list) and colons (ranges).
Private Function NormalizeListInput(ByVal s As String) As String
    s = CStr(s)

    ' Remove common hidden characters
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr$(160), "") ' non-breaking space

    ' Normalise all "list separators" to comma
    s = Replace(s, ";", ",")
    s = Replace(s, " ", ",")

    ' Normalise all "range separators" to colon
    s = Replace(s, "-", ":")
    s = Replace(s, ChrW(8211), ":") ' FIX: ChrW() required for Unicode U+2013 en dash; Chr$() is ANSI-only (0-255)
    s = Replace(s, ChrW(8212), ":") ' FIX: ChrW() required for Unicode U+2014 em dash; Chr$() is ANSI-only (0-255)
    s = Replace(s, ".", ":")

    ' Collapse runs of commas produced by the above
    Do While InStr(s, ",,") > 0
        s = Replace(s, ",,", ",")
    Loop

    ' Trim leading/trailing commas
    If Left$(s, 1) = "," Then s = Mid$(s, 2)
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)

    NormalizeListInput = Trim$(s)
End Function

' Parse a row specification string into an array of row numbers.
' Supports: ranges (2:500), lists (2,5,10), and mixed (2,5,10:15).
' Returns a single-element array containing 0 when input is empty/invalid.
Public Function ParseRowList(ByVal inputText As String) As Long()
    Dim parts() As String
    Dim i As Long
    Dim tmpList As Collection
    Dim token As String
    Dim subParts() As String
    Dim startRow As Long, endRow As Long
    Dim r As Long
    Dim arr() As Long

    Set tmpList = New Collection

    inputText = NormalizeListInput(inputText)
    If inputText = "" Then GoTo EmptyExit

    parts = Split(inputText, ",")

    For i = LBound(parts) To UBound(parts)
        token = parts(i)

        If token <> vbNullString Then
            If InStr(1, token, ":", vbBinaryCompare) > 0 Then
                subParts = Split(token, ":")

                If UBound(subParts) >= 1 Then
                    If IsNumeric(subParts(0)) And IsNumeric(subParts(1)) Then
                        startRow = CLng(subParts(0))
                        endRow = CLng(subParts(1))

                        If endRow < startRow Then
                            r = startRow
                            startRow = endRow
                            endRow = r
                        End If

                        For r = startRow To endRow
                            tmpList.Add r
                        Next r
                    End If
                End If
            Else
                If IsNumeric(token) Then
                    tmpList.Add CLng(token)
                End If
            End If
        End If
    Next i

    If tmpList.Count = 0 Then GoTo EmptyExit

    ReDim arr(0 To tmpList.Count - 1)
    For i = 1 To tmpList.Count
        arr(i - 1) = CLng(tmpList(i))
    Next i

    ParseRowList = arr
    Exit Function

EmptyExit:
    ReDim arr(0 To 0)
    arr(0) = 0
    ParseRowList = arr
End Function

' Parse a column specification string into an array of column numbers.
' Supports letter ranges (H:AK), lists (H,J,M), and mixed (H,J:K,N).
' Returns a single-element array containing 0 when input is empty/invalid.
Public Function ParseColumnList(ByVal inputText As String) As Long()
    Dim parts() As String
    Dim i As Long
    Dim tmpList As Collection
    Dim token As String
    Dim subParts() As String
    Dim startCol As Long, endCol As Long
    Dim c As Long
    Dim arr() As Long

    Set tmpList = New Collection

    Debug.Print "Raw clauseColsText: [" & inputText & "]"
    inputText = NormalizeListInput(inputText)
    Debug.Print "Normalized clauseColsText: [" & inputText & "]"

    If inputText = "" Then GoTo EmptyExit

    parts = Split(inputText, ",")
    Debug.Print "Parts count: " & (UBound(parts) - LBound(parts) + 1)

    For i = LBound(parts) To UBound(parts)
        token = parts(i)
        Debug.Print "  Token " & i & ": [" & token & "]"

        If token <> vbNullString Then
            If InStr(1, token, ":", vbBinaryCompare) > 0 Then
                subParts = Split(token, ":")

                If UBound(subParts) >= 1 Then
                    Debug.Print "    Range from [" & subParts(0) & "] to [" & subParts(1) & "]"
                    startCol = ColumnLetterToNumber(subParts(0))
                    endCol = ColumnLetterToNumber(subParts(1))
                    Debug.Print "    startCol=" & startCol & " endCol=" & endCol

                    If startCol > 0 And endCol > 0 Then
                        If endCol < startCol Then
                            c = startCol
                            startCol = endCol
                            endCol = c
                        End If

                        For c = startCol To endCol
                            tmpList.Add c
                            Debug.Print "      Added col " & c
                        Next c
                    Else
                        Debug.Print "    Skipping: invalid start/end"
                    End If
                End If
            Else
                startCol = ColumnLetterToNumber(token)
                Debug.Print "    Single col token; ColumnLetterToNumber=" & startCol
                If startCol > 0 Then
                    tmpList.Add startCol
                    Debug.Print "      Added col " & startCol
                Else
                    Debug.Print "    Skipping: invalid single col token"
                End If
            End If
        End If
    Next i

    If tmpList.Count = 0 Then
        Debug.Print "No valid cols found; returning [0]"
        GoTo EmptyExit
    End If

    ReDim arr(0 To tmpList.Count - 1)
    For i = 1 To tmpList.Count
        arr(i - 1) = CLng(tmpList(i))
    Next i
    ParseColumnList = arr
    Exit Function

EmptyExit:
    ReDim arr(0 To 0)
    arr(0) = 0
    ParseColumnList = arr
End Function

' Percent-encode a string for use in a mailto: hyperlink.
Public Function URLEncode(ByVal sText As String) As String
    Dim i As Long
    Dim ch As String
    Dim sOut As String

    For i = 1 To Len(sText)
        ch = Mid$(sText, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122   ' 0-9 A-Z a-z
                sOut = sOut & ch
            Case 32
                sOut = sOut & "%20"
            Case Else
                sOut = sOut & "%" & Right$("0" & Hex(AscW(ch) And &HFF), 2)
        End Select
    Next i

    URLEncode = sOut
End Function

' Normalise a single column letter (e.g. " al " -> "AL").
Public Function NormaliseCol(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " ", "")
    NormaliseCol = s
End Function

' Normalise a comma-separated column list (e.g. " h : ak " -> "H:AK").
Public Function NormaliseCols(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " ", "")
    NormaliseCols = s
End Function

' Normalise a row-ranges string (e.g. " 2 - 500 " -> "2-500").
Public Function NormaliseRowRanges(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, " ", "")
    NormaliseRowRanges = s
End Function

' Validate the three normalised inputs.
' Returns "" if all are valid; otherwise returns a user-readable error message.
Public Function ValidateInputs(ByVal defsCol As String, _
                               ByVal clauseCols As String, _
                               ByVal rowRanges As String) As String
    If defsCol = "" Then
        ValidateInputs = "Definitions column cannot be empty (e.g. AL)."
        Exit Function
    End If

    If ColumnLetterToNumber(defsCol) = 0 Then
        ValidateInputs = "Definitions column is not a valid column letter (e.g. AL): """ & defsCol & """"
        Exit Function
    End If

    If clauseCols = "" Then
        ValidateInputs = "Clause columns cannot be empty (e.g. H:AK or C,E,G)."
        Exit Function
    End If

    Dim arrCols() As Long
    arrCols = ParseColumnList(clauseCols)
    If UBound(arrCols) = 0 And arrCols(0) = 0 Then
        ValidateInputs = "Clause columns could not be parsed (e.g. H:AK or C,E,G): """ & clauseCols & """"
        Exit Function
    End If

    If rowRanges = "" Then
        ValidateInputs = "Row ranges cannot be empty (e.g. 2:500 or 2,5,10)."
        Exit Function
    End If

    Dim arrRows() As Long
    arrRows = ParseRowList(rowRanges)
    If UBound(arrRows) = 0 And arrRows(0) = 0 Then
        ValidateInputs = "Row ranges could not be parsed (e.g. 2:500 or 2,5,10): """ & rowRanges & """"
        Exit Function
    End If

    ValidateInputs = ""
End Function

' Dump character-by-character breakdown of a string to the Immediate Window.
' Useful for diagnosing invisible / non-printing characters.
Public Sub DebugShowChars(ByVal label As String, ByVal s As String)
    Dim i As Long, ch As String
    Debug.Print "---- " & label & " ----"
    Debug.Print "Len=" & Len(s) & "  [" & s & "]"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Debug.Print i & ": '" & ch & "'  AscW=" & AscW(ch)
    Next i
End Sub
