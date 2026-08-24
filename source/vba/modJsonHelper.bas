Attribute VB_Name = "modJsonHelper"
Option Explicit

#If Mac Then

    ' Core Foundation:
    ' CFCharacterSetGetPredefined
    Private Declare PtrSafe Function CFCharacterSetGetPredefined Lib _
        "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" ( _
            ByVal theSetIdentifier As Long _
        ) As LongPtr

    ' Boolean CFCharacterSetIsLongCharacterMember(CFCharacterSetRef, UTF-32 code point)
    Private Declare PtrSafe Function CFCharacterSetIsLongCharacterMember Lib _
        "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" ( _
            ByVal theSet As LongPtr, _
            ByVal theChar As Long _
        ) As Boolean

    Private Const kCFCharacterSetDecimalDigit As Long = 4
    Private Const kCFCharacterSetLetter As Long = 5

#elseif VBA7 Then
    Private Declare PtrSafe Function GetStringTypeW Lib "kernel32" ( _
        ByVal dwInfoType As Long, _
        ByVal lpSrcStr As LongPtr, _
        ByVal cchSrc As Long, _
        ByRef lpCharType As Integer _
    ) As Long
#Else
    Private Declare Function GetStringTypeW Lib "kernel32" ( _
        ByVal dwInfoType As Long, _
        ByVal lpSrcStr As Long, _
        ByVal cchSrc As Long, _
        ByRef lpCharType As Integer _
    ) As Long
#End If

Private Const CT_CTYPE1 As Long = 1
Private Const C1_DIGIT As Integer = &H4
Private Const C1_ALPHA As Integer = &H100

Public Function IsUnicodeLetterOrDigit(ByVal ch As String) As Boolean

#If Mac Then
    IsUnicodeLetterOrDigit = IsUnicodeLetterOrDigitMac(ch)
#Else
    IsUnicodeLetterOrDigit = IsUnicodeLetterOrDigitWindows(ch)
#End If

End Function

#If Mac Then

Private Function IsUnicodeLetterOrDigitMac(ByVal ch As String) As Boolean
    Dim letterSet As LongPtr
    Dim digitSet As LongPtr
    Dim codePoint As Long

    If Len(ch) = 0 Then Exit Function

    ' On current Office for Mac, use Asc rather than AscW.
    ' It is reliable for the Mac VBA character representation, unlike AscW.
    codePoint = Asc(Left$(ch, 1))

    letterSet = CFCharacterSetGetPredefined(kCFCharacterSetLetter)
    digitSet = CFCharacterSetGetPredefined(kCFCharacterSetDecimalDigit)

    IsUnicodeLetterOrDigitMac = _
        CBool(CFCharacterSetIsLongCharacterMember(letterSet, codePoint)) _
        Or CBool(CFCharacterSetIsLongCharacterMember(digitSet, codePoint))
End Function

#else

Private Function IsUnicodeLetterOrDigitWindows(ByVal ch As String) As Boolean
    Dim charType As Integer

    If Len(ch) = 0 Then
        IsUnicodeLetterOrDigitWindows = False
        Exit Function
    End If

    If GetStringTypeW(CT_CTYPE1, StrPtr(ch), 1, charType) = 0 Then
        IsUnicodeLetterOrDigitWindows = False
        Exit Function
    End If

    IsUnicodeLetterOrDigitWindows = _
        ((charType And C1_ALPHA) <> 0) _
        Or ((charType And C1_DIGIT) <> 0)
End Function

#end if

Public Function IsWhitespaceOnly(ByVal ch As String) As Boolean
    IsWhitespaceOnly = (ch = " " Or ch = vbTab)
End Function

Public Function IsWhitespaceChar(ByVal ch As String) As Boolean
    IsWhitespaceChar = ( _
        ch = " " _
        Or ch = vbTab _
        Or ch = vbCr _
        Or ch = vbLf _
    )
End Function

Public Function IsTokenPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", ")", "[", "]", "{", "}", _
             ".", ",", ";", ":", "!", "?", _
             """", "'", "+", "=", "*", "&", "|", "\", "/", "<", ">", _
             "@", "#", "$", "%", "^", "_", "~", "`", "-"
            IsTokenPunctuation = True

        Case ChrW(&H2018), ChrW(&H2019), ChrW(&H201A), ChrW(&H201B), _
             ChrW(&H201C), ChrW(&H201D), ChrW(&H201E), ChrW(&H201F), _
             ChrW(&HAB), ChrW(&HBB), ChrW(&H2039), ChrW(&H203A), _
             ChrW(&H2013), ChrW(&H2014), ChrW(&H2026)
            IsTokenPunctuation = True

        Case Else
            IsTokenPunctuation = False
    End Select
End Function

Public Function IsWordChar(ByVal ch As String) As Boolean
    Dim normalized As String

    normalized = NormalizeAnalysisText(ch)

    ' Characters deliberately removed by normalization remain part
    ' of the raw token, preserving correspondence with the normalized form.
    If normalized = "" Then
        IsWordChar = True
        Exit Function
    End If

    ' Deliberate word-internal connectors.
    If normalized = "-" _
       Or normalized = "/" _
       Or normalized = "'" Then

        IsWordChar = True
        Exit Function
    End If

    IsWordChar = IsUnicodeLetterOrDigit(normalized)
End Function

Public Function ParseJsonStringArray(ByVal s As String) As Collection
    Dim c As New Collection
    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim current As String
    
    inString = False
    escaped = False
    current = ""
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        
        If inString Then
            If escaped Then
                current = current & "\" & ch
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                c.Add JsonUnescape(current)
                current = ""
                inString = False
            Else
                current = current & ch
            End If
        Else
            If ch = """" Then
                inString = True
                current = ""
            ElseIf ch = "[" Or ch = "]" Or ch = "{" Or ch = "}" Then
                Set ParseJsonStringArray = Nothing
                Exit Function
            End If
        End If
    Next i
    
    If inString Or escaped Then
        Set ParseJsonStringArray = Nothing
        Exit Function
    End If
    
    Set ParseJsonStringArray = c
End Function

Public Function FindMatchingBracket(ByVal s As String, ByVal openPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim ch As String
    
    depth = 0
    inString = False
    escaped = False
    
    For i = openPos To Len(s)
        ch = Mid$(s, i, 1)
        
        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            If ch = """" Then
                inString = True
            ElseIf ch = "[" Then
                depth = depth + 1
            ElseIf ch = "]" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingBracket = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindMatchingBracket = 0
End Function

Public Function GetParagraphTextWithoutMark(ByVal rng As Range) As String
    Dim s As String
    s = rng.Text
    If Len(s) > 0 Then
        If Right$(s, 1) = vbCr Then
            s = Left$(s, Len(s) - 1)
        End If
    End If
    GetParagraphTextWithoutMark = s
End Function

Public Function IsSubstantiveParagraph(ByVal s As String) As Boolean
    Dim t As String
    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Trim$(t)
    IsSubstantiveParagraph = (t <> "")
End Function

Public Function NormalizeParagraphText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    
    Do While Len(s) > 0 And Right$(s, 1) = vbLf
        s = Left$(s, Len(s) - 1)
    Loop
    
    NormalizeParagraphText = s
End Function

Public Function ExtractAssistantContent(ByVal jsonText As String) As String
    Dim pChoices As Long
    Dim pMsg As Long
    Dim pContent As Long
    Dim pValue As Long
    Dim qValue As Long
    
    pChoices = InStr(1, jsonText, """choices""", vbTextCompare)
    If pChoices = 0 Then Exit Function
    
    pMsg = InStr(pChoices, jsonText, """message""", vbTextCompare)
    If pMsg = 0 Then Exit Function
    
    pContent = InStr(pMsg, jsonText, """content""", vbTextCompare)
    If pContent = 0 Then Exit Function
    
    ' Was: InStr(pContent + 9, jsonText, """) - an unterminated string literal.
    pValue = InStr(pContent + 9, jsonText, Chr$(34))
    If pValue = 0 Then Exit Function
    pValue = pValue + 1
    
    qValue = FindJsonStringEnd(jsonText, pValue)
    If qValue = 0 Then Exit Function
    
    ExtractAssistantContent = JsonUnescape(Mid$(jsonText, pValue, qValue - pValue))
End Function

Public Function FindJsonStringEnd(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean
    
    escaped = False
    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If escaped Then
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            FindJsonStringEnd = i
            Exit Function
        End If
    Next i
    
    FindJsonStringEnd = 0
End Function

Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        
        Select Case code
            Case 34
                result = result & "\"""     ' backslash + quote
            Case 92
                result = result & "\\"
            Case 8
                result = result & "\b"
            Case 9
                result = result & "\t"
            Case 10
                result = result & "\n"
            Case 12
                result = result & "\f"
            Case 13
                result = result & "\r"
            Case 0 To 31
                result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else
                result = result & ch
        End Select
    Next i
    
    JsonEscape = result
End Function

Public Function JsonUnescape(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean
    Dim result As String
    Dim hex4 As String
    
    escaped = False
    result = ""
    
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        
        If escaped Then
            Select Case ch
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case """"
                    result = result & """"    ' a single quote character
                Case "\"
                    result = result & "\"
                Case "u"
                    If i + 4 <= Len(s) Then
                        hex4 = Mid$(s, i + 1, 4)
                        If IsHex4(hex4) Then
                            result = result & ChrW$(CLng("&H" & hex4))
                            i = i + 4
                        Else
                            result = result & "\u"
                        End If
                    Else
                        result = result & "\u"
                    End If
                Case Else
                    result = result & "\" & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        Else
            result = result & ch
        End If
        
        i = i + 1
    Loop
    
    If escaped Then
        result = result & "\"
    End If
    
    JsonUnescape = result
End Function

Public Function IsHex4(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    If Len(s) <> 4 Then Exit Function
    
    For i = 1 To 4
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789ABCDEFabcdef", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i
    
    IsHex4 = True
End Function

Public Function CleanupModelOutput(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Trim$(s)

    ' Reasoning models that do not use a separate reasoning_content field emit
    ' their chain of thought inline; it must not reach the JSON parser.
    s = StripThinkBlocks(s)

    If Left$(s, 3) = "```" Then
        s = StripCodeFences(s)
    End If
    
    CleanupModelOutput = Trim$(s)
End Function

' Removes <think> ... </think> blocks, including an unterminated leading one.
Public Function StripThinkBlocks(ByVal s As String) As String
    Dim pOpen As Long
    Dim pClose As Long

    Do
        pOpen = InStr(1, s, "<think>", vbTextCompare)
        If pOpen = 0 Then Exit Do

        pClose = InStr(pOpen, s, "</think>", vbTextCompare)
        If pClose = 0 Then
            ' No closing tag: everything from the tag on is reasoning.
            s = Left$(s, pOpen - 1)
            Exit Do
        End If

        s = Left$(s, pOpen - 1) & Mid$(s, pClose + 8)
    Loop

    StripThinkBlocks = Trim$(s)
End Function

Public Function StripCodeFences(ByVal s As String) As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim t As String
    
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    lines = Split(s, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        t = Trim$(lines(i))
        If Left$(t, 3) <> "```" Then
            If result = "" Then
                result = lines(i)
            Else
                result = result & vbLf & lines(i)
            End If
        End If
    Next i
    
    StripCodeFences = Trim$(result)
End Function

Public Function IsAllDigitsText(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    If Len(s) = 0 Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    
    IsAllDigitsText = True
End Function


' This is for debugging
Public Function CharCodes(ByVal s As String) As String
    Dim i As Long
    Dim result As String

    For i = 1 To Len(s)
        If Len(result) > 0 Then result = result & " "
        result = result & "U+" & Right$("0000" & Hex$(AscW(Mid$(s, i, 1))), 4)
    Next i

    CharCodes = result
End Function


Public Sub ShowSelectedCharCodes()
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim msg As String

    s = Selection.Text

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        msg = msg & "'" & ch & "'  Hex=" & Hex$(AscW(ch)) & "  Dec=" & AscW(ch) & vbCrLf
    Next i

    MsgBox msg, vbInformation, "Selected character codes"
End Sub

Public Function NormalizeAnalysisText(ByVal s As String) As String
    ' Dash-like characters -> standard hyphen
    s = Replace(s, ChrW(&H2010), "-") ' hyphen
    s = Replace(s, ChrW(&H2011), "-") ' non-breaking hyphen
    s = Replace(s, ChrW(&H2012), "-") ' figure dash
    s = Replace(s, ChrW(&H2013), "-") ' en dash
    s = Replace(s, ChrW(&H2014), "-") ' em dash
    s = Replace(s, ChrW(&H2015), "-") ' horizontal bar
    s = Replace(s, ChrW(&H2212), "-") ' minus sign
    ' Additional dash-like characters that may appear in Word and should become "-"
    s = Replace(s, ChrW(&H2043), "-")   ' hyphen bullet
    s = Replace(s, ChrW(&HFF0D), "-")   ' fullwidth hyphen-minus
        
    ' RS (0x1E) used by some chatbots as weird hyphen-like separator -> treat as hyphen
    s = Replace(s, ChrW(&H1E), "-")
    s = Replace(s, ChrW(&HFE63), "-")  ' small hyphen-minus, also used by some LLMs as hypen separator
    
    ' Apostrophe-like characters -> straight apostrophe
    s = Replace(s, ChrW(&H2018), "'") ' left single quotation mark
    s = Replace(s, ChrW(&H2019), "'") ' right single quotation mark / apostrophe
    s = Replace(s, ChrW(&H201B), "'") ' single high-reversed-9 quotation mark
    s = Replace(s, ChrW(&H2032), "'") ' prime
    s = Replace(s, ChrW(&HB4), "'")   ' acute accent often used as apostrophe
    
    ' Space-like characters -> normal space
    s = Replace(s, ChrW(&HA0), " ")   ' no-break space
    s = Replace(s, ChrW(&H2000), " ") ' en quad
    s = Replace(s, ChrW(&H2001), " ") ' em quad
    s = Replace(s, ChrW(&H2002), " ") ' en space
    s = Replace(s, ChrW(&H2003), " ") ' em space
    s = Replace(s, ChrW(&H2004), " ") ' three-per-em space
    s = Replace(s, ChrW(&H2005), " ") ' four-per-em space
    s = Replace(s, ChrW(&H2006), " ") ' six-per-em space
    s = Replace(s, ChrW(&H2007), " ") ' figure space
    s = Replace(s, ChrW(&H2008), " ") ' punctuation space
    s = Replace(s, ChrW(&H2009), " ") ' thin space
    s = Replace(s, ChrW(&H200A), " ") ' hair space
    s = Replace(s, ChrW(&H202F), " ") ' narrow no-break space
    s = Replace(s, ChrW(&H205F), " ") ' medium mathematical space
    s = Replace(s, ChrW(&H3000), " ") ' ideographic space
    
    ' Remove invisible format/control chars that should not affect matching
    s = Replace(s, ChrW(&HAD), "")    ' soft hyphen
    s = Replace(s, ChrW(&H200B), "")  ' zero width space
    s = Replace(s, ChrW(&H200C), "")  ' zero width non-joiner
    s = Replace(s, ChrW(&H200D), "")  ' zero width joiner
    s = Replace(s, ChrW(&H2060), "")  ' word joiner
    s = Replace(s, ChrW(&HFEFF), "")  ' zero width no-break space / BOM
    
    NormalizeAnalysisText = s
End Function
