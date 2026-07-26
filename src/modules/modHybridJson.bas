Option Explicit

' Minimal strict JSON support for the local hybrid protocol. The parser accepts
' JSON objects/arrays/primitives, rejects duplicate object keys and trailing
' content, and never evaluates script.

Private mHybridJsonText As String
Private mHybridJsonPosition As Long

Public Function HybridJsonParseObject(ByVal jsonText As String) As Object
    mHybridJsonText = jsonText
    mHybridJsonPosition = 1
    HybridJsonSkipWhitespace
    If HybridJsonPeek() <> "{" Then Err.Raise 5, "HybridJsonParseObject", "The JSON root must be an object."
    Set HybridJsonParseObject = HybridJsonReadObject()
    HybridJsonSkipWhitespace
    If mHybridJsonPosition <= Len(mHybridJsonText) Then Err.Raise 5, "HybridJsonParseObject", "Unexpected content follows the JSON value."
End Function

Public Function HybridJsonEscape(ByVal value As String) As String
    Dim i As Long
    Dim codeUnit As Long
    Dim ch As String
    Dim result As String

    result = """"
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        codeUnit = AscW(ch)
        If codeUnit < 0 Then codeUnit = codeUnit + &H10000
        Select Case ch
            Case """": result = result & "\"""
            Case "\": result = result & "\\"
            Case vbBack: result = result & "\b"
            Case vbFormFeed: result = result & "\f"
            Case vbLf: result = result & "\n"
            Case vbCr: result = result & "\r"
            Case vbTab: result = result & "\t"
            Case Else
                If codeUnit < 32 Then
                    result = result & "\u" & Right$("0000" & Hex$(codeUnit), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    HybridJsonEscape = result & """"
End Function

Private Function HybridJsonReadObject() As Object
    Dim result As Object
    Dim key As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare
    HybridJsonExpect "{"
    HybridJsonSkipWhitespace
    If HybridJsonPeek() = "}" Then
        mHybridJsonPosition = mHybridJsonPosition + 1
        Set HybridJsonReadObject = result
        Exit Function
    End If

    Do
        HybridJsonSkipWhitespace
        If HybridJsonPeek() <> """" Then Err.Raise 5, "HybridJsonReadObject", "A JSON object key must be a string."
        key = HybridJsonReadString()
        If result.Exists(key) Then Err.Raise 5, "HybridJsonReadObject", "A JSON object contains a duplicate key."
        HybridJsonSkipWhitespace
        HybridJsonExpect ":"
        HybridJsonSkipWhitespace
        HybridJsonAddDictionaryValue result, key
        HybridJsonSkipWhitespace
        Select Case HybridJsonPeek()
            Case "}"
                mHybridJsonPosition = mHybridJsonPosition + 1
                Exit Do
            Case ","
                mHybridJsonPosition = mHybridJsonPosition + 1
            Case Else
                Err.Raise 5, "HybridJsonReadObject", "A JSON object is missing a comma or closing brace."
        End Select
    Loop

    Set HybridJsonReadObject = result
End Function

Private Function HybridJsonReadArray() As Collection
    Dim result As Collection
    Set result = New Collection

    HybridJsonExpect "["
    HybridJsonSkipWhitespace
    If HybridJsonPeek() = "]" Then
        mHybridJsonPosition = mHybridJsonPosition + 1
        Set HybridJsonReadArray = result
        Exit Function
    End If

    Do
        HybridJsonSkipWhitespace
        HybridJsonAddCollectionValue result
        HybridJsonSkipWhitespace
        Select Case HybridJsonPeek()
            Case "]"
                mHybridJsonPosition = mHybridJsonPosition + 1
                Exit Do
            Case ","
                mHybridJsonPosition = mHybridJsonPosition + 1
            Case Else
                Err.Raise 5, "HybridJsonReadArray", "A JSON array is missing a comma or closing bracket."
        End Select
    Loop

    Set HybridJsonReadArray = result
End Function

Private Sub HybridJsonAddDictionaryValue(ByVal target As Object, ByVal key As String)
    Dim marker As String
    Dim child As Object
    marker = HybridJsonPeek()

    If marker = "{" Then
        Set child = HybridJsonReadObject()
        target.Add key, child
    ElseIf marker = "[" Then
        Set child = HybridJsonReadArray()
        target.Add key, child
    Else
        target.Add key, HybridJsonReadPrimitive()
    End If
End Sub

Private Sub HybridJsonAddCollectionValue(ByVal target As Collection)
    Dim marker As String
    Dim child As Object
    marker = HybridJsonPeek()

    If marker = "{" Then
        Set child = HybridJsonReadObject()
        target.Add child
    ElseIf marker = "[" Then
        Set child = HybridJsonReadArray()
        target.Add child
    Else
        target.Add HybridJsonReadPrimitive()
    End If
End Sub

Private Function HybridJsonReadPrimitive() As Variant
    Select Case HybridJsonPeek()
        Case """"
            HybridJsonReadPrimitive = HybridJsonReadString()
        Case "t"
            HybridJsonExpectLiteral "true"
            HybridJsonReadPrimitive = True
        Case "f"
            HybridJsonExpectLiteral "false"
            HybridJsonReadPrimitive = False
        Case "n"
            HybridJsonExpectLiteral "null"
            HybridJsonReadPrimitive = Null
        Case "-", "0" To "9"
            HybridJsonReadPrimitive = HybridJsonReadNumber()
        Case Else
            Err.Raise 5, "HybridJsonReadPrimitive", "The JSON value is invalid."
    End Select
End Function

Private Function HybridJsonReadString() As String
    Dim result As String
    Dim ch As String
    Dim escaped As String
    Dim codeUnit As Long
    Dim lowCodeUnit As Long
    Dim nextCharacter As String

    HybridJsonExpect """"
    Do While mHybridJsonPosition <= Len(mHybridJsonText)
        ch = Mid$(mHybridJsonText, mHybridJsonPosition, 1)
        mHybridJsonPosition = mHybridJsonPosition + 1
        If ch = """" Then
            HybridJsonReadString = result
            Exit Function
        End If
        codeUnit = HybridJsonUnsignedCodeUnit(ch)
        If codeUnit < 32 Then Err.Raise 5, "HybridJsonReadString", "A JSON string contains an unescaped control character."
        If ch <> "\" Then
            If codeUnit >= 55296 And codeUnit <= 56319 Then
                If mHybridJsonPosition > Len(mHybridJsonText) Then Err.Raise 5, "HybridJsonReadString", "A JSON string contains an unpaired high surrogate."
                nextCharacter = Mid$(mHybridJsonText, mHybridJsonPosition, 1)
                lowCodeUnit = HybridJsonUnsignedCodeUnit(nextCharacter)
                If lowCodeUnit < 56320 Or lowCodeUnit > 57343 Then Err.Raise 5, "HybridJsonReadString", "A JSON string contains an unpaired high surrogate."
                result = result & ch & nextCharacter
                mHybridJsonPosition = mHybridJsonPosition + 1
            ElseIf codeUnit >= 56320 And codeUnit <= 57343 Then
                Err.Raise 5, "HybridJsonReadString", "A JSON string contains an unpaired low surrogate."
            Else
                result = result & ch
            End If
        Else
            If mHybridJsonPosition > Len(mHybridJsonText) Then Err.Raise 5, "HybridJsonReadString", "A JSON escape is incomplete."
            escaped = Mid$(mHybridJsonText, mHybridJsonPosition, 1)
            mHybridJsonPosition = mHybridJsonPosition + 1
            Select Case escaped
                Case """", "\", "/": result = result & escaped
                Case "b": result = result & vbBack
                Case "f": result = result & vbFormFeed
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    codeUnit = HybridJsonReadHexCodeUnit()
                    If codeUnit >= 55296 And codeUnit <= 56319 Then
                        If Mid$(mHybridJsonText, mHybridJsonPosition, 2) <> "\u" Then Err.Raise 5, "HybridJsonReadString", "A Unicode escape contains an unpaired high surrogate."
                        mHybridJsonPosition = mHybridJsonPosition + 2
                        lowCodeUnit = HybridJsonReadHexCodeUnit()
                        If lowCodeUnit < 56320 Or lowCodeUnit > 57343 Then Err.Raise 5, "HybridJsonReadString", "A Unicode escape contains an unpaired high surrogate."
                        result = result & HybridJsonCodeUnit(codeUnit) & HybridJsonCodeUnit(lowCodeUnit)
                    ElseIf codeUnit >= 56320 And codeUnit <= 57343 Then
                        Err.Raise 5, "HybridJsonReadString", "A Unicode escape contains an unpaired low surrogate."
                    Else
                        result = result & HybridJsonCodeUnit(codeUnit)
                    End If
                Case Else
                    Err.Raise 5, "HybridJsonReadString", "A JSON string contains an invalid escape."
            End Select
        End If
    Loop
    Err.Raise 5, "HybridJsonReadString", "A JSON string is unterminated."
End Function

Private Function HybridJsonUnsignedCodeUnit(ByVal character As String) As Long
    HybridJsonUnsignedCodeUnit = AscW(character)
    If HybridJsonUnsignedCodeUnit < 0 Then HybridJsonUnsignedCodeUnit = HybridJsonUnsignedCodeUnit + &H10000
End Function

Private Function HybridJsonReadHexCodeUnit() As Long
    Dim hexText As String
    Dim i As Long
    Dim ch As String

    If mHybridJsonPosition + 3 > Len(mHybridJsonText) Then Err.Raise 5, "HybridJsonReadHexCodeUnit", "A Unicode escape is incomplete."
    hexText = Mid$(mHybridJsonText, mHybridJsonPosition, 4)
    For i = 1 To 4
        ch = Mid$(hexText, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Err.Raise 5, "HybridJsonReadHexCodeUnit", "A Unicode escape is invalid."
    Next i
    mHybridJsonPosition = mHybridJsonPosition + 4
    HybridJsonReadHexCodeUnit = CLng("&H" & hexText)
    If HybridJsonReadHexCodeUnit < 0 Then HybridJsonReadHexCodeUnit = HybridJsonReadHexCodeUnit + &H10000
End Function

Private Function HybridJsonCodeUnit(ByVal codeUnit As Long) As String
    If codeUnit > &H7FFF Then
        HybridJsonCodeUnit = ChrW$(codeUnit - &H10000)
    Else
        HybridJsonCodeUnit = ChrW$(codeUnit)
    End If
End Function

Private Function HybridJsonReadNumber() As Variant
    Dim startPosition As Long
    Dim token As String
    Dim ch As String
    Dim isFloating As Boolean

    startPosition = mHybridJsonPosition
    If HybridJsonPeek() = "-" Then mHybridJsonPosition = mHybridJsonPosition + 1
    If HybridJsonPeek() = "0" Then
        mHybridJsonPosition = mHybridJsonPosition + 1
        If HybridJsonPeek() Like "[0-9]" Then Err.Raise 5, "HybridJsonReadNumber", "A JSON number has a leading zero."
    Else
        HybridJsonRequireDigit
        Do While HybridJsonPeek() Like "[0-9]"
            mHybridJsonPosition = mHybridJsonPosition + 1
        Loop
    End If
    If HybridJsonPeek() = "." Then
        isFloating = True
        mHybridJsonPosition = mHybridJsonPosition + 1
        HybridJsonRequireDigit
        Do While HybridJsonPeek() Like "[0-9]"
            mHybridJsonPosition = mHybridJsonPosition + 1
        Loop
    End If
    ch = HybridJsonPeek()
    If ch = "e" Or ch = "E" Then
        isFloating = True
        mHybridJsonPosition = mHybridJsonPosition + 1
        ch = HybridJsonPeek()
        If ch = "+" Or ch = "-" Then mHybridJsonPosition = mHybridJsonPosition + 1
        HybridJsonRequireDigit
        Do While HybridJsonPeek() Like "[0-9]"
            mHybridJsonPosition = mHybridJsonPosition + 1
        Loop
    End If

    token = Mid$(mHybridJsonText, startPosition, mHybridJsonPosition - startPosition)
    If isFloating Then
        HybridJsonReadNumber = CDbl(token)
    Else
        HybridJsonReadNumber = CLng(token)
    End If
End Function

Private Sub HybridJsonRequireDigit()
    If Not (HybridJsonPeek() Like "[0-9]") Then Err.Raise 5, "HybridJsonRequireDigit", "A JSON number is incomplete."
End Sub

Private Sub HybridJsonExpectLiteral(ByVal literal As String)
    If Mid$(mHybridJsonText, mHybridJsonPosition, Len(literal)) <> literal Then Err.Raise 5, "HybridJsonExpectLiteral", "A JSON literal is invalid."
    mHybridJsonPosition = mHybridJsonPosition + Len(literal)
End Sub

Private Sub HybridJsonExpect(ByVal expected As String)
    If HybridJsonPeek() <> expected Then Err.Raise 5, "HybridJsonExpect", "Unexpected JSON character."
    mHybridJsonPosition = mHybridJsonPosition + 1
End Sub

Private Sub HybridJsonSkipWhitespace()
    Dim ch As String
    Do While mHybridJsonPosition <= Len(mHybridJsonText)
        ch = Mid$(mHybridJsonText, mHybridJsonPosition, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        mHybridJsonPosition = mHybridJsonPosition + 1
    Loop
End Sub

Private Function HybridJsonPeek() As String
    If mHybridJsonPosition <= Len(mHybridJsonText) Then HybridJsonPeek = Mid$(mHybridJsonText, mHybridJsonPosition, 1)
End Function
