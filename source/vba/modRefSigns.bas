Attribute VB_Name = "modRefSigns"
Option Explicit

#If VBA7 Then
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

Private Function IsUnicodeLetterOrDigit(ByVal ch As String) As Boolean
    Dim charType As Integer

    If Len(ch) = 0 Then
        IsUnicodeLetterOrDigit = False
        Exit Function
    End If

    If GetStringTypeW(CT_CTYPE1, StrPtr(ch), 1, charType) = 0 Then
        IsUnicodeLetterOrDigit = False
        Exit Function
    End If

    IsUnicodeLetterOrDigit = _
        ((charType And C1_ALPHA) <> 0) _
        Or ((charType And C1_DIGIT) <> 0)
End Function


Public Sub Patent_Tools_Settings()
    frmPatentToolsSettings.Show
End Sub

Public Sub Insert_Reference_Signs()
    Dim refTable As String
    Dim targetRange As Range
    Dim paraRanges As Collection
    Dim paraTexts As Collection
    Dim p As Paragraph
    Dim paraText As String
    Dim promptText As String
    Dim endpoint As String
    Dim requestJson As String
    Dim responseJson As String
    Dim assistantText As String
    Dim rewrittenParas As Collection
    Dim http As Object
    Dim i As Long
    Dim useSelection As Boolean
    Dim workRng As Range
    
    Dim targetDoc As Document
    Dim targetWindow As Window
    Dim oldTrackRevisions As Boolean
    Dim failedParagraphs As String
    Dim okInsert As Boolean
    
    Rem Populate the g... global variables
    LoadPatentToolsSettings
   
    refTable = GetReferenceListFromForm()
    If Trim$(refTable) = "" Then
        MsgBox "Reference-sign table is required.", vbExclamation
        Exit Sub
    End If
    
    useSelection = False
    If Not Selection Is Nothing Then
        If Len(Trim$(Replace(Selection.Range.Text, vbCr, ""))) > 0 Then
            useSelection = True
            Set targetRange = Selection.Range.Duplicate
        End If
    End If
    
    If Not useSelection Then
        If MsgBox( _
            "Claims are not selected (highlighted)." & vbCrLf & vbCrLf & _
            "Full document processing is costly." & vbCrLf & vbCrLf & _
            "Proceed anyway?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "PatentTools" _
        ) <> vbYes Then
            Exit Sub
        End If
        Set targetRange = ActiveDocument.Content
    End If
    
    
    Set targetDoc = targetRange.Document
    Set targetWindow = targetDoc.ActiveWindow
    
    Set paraRanges = New Collection
    Set paraTexts = New Collection
    
    For Each p In targetRange.Paragraphs
        paraText = GetParagraphTextWithoutMark(p.Range.Duplicate)
        If IsSubstantiveParagraph(paraText) Then
            paraRanges.Add p.Range.Duplicate
            paraTexts.Add paraText
        End If
    Next p
    
    If paraTexts.Count = 0 Then
        MsgBox "No substantive paragraphs found in the current selection/document.", vbExclamation
        Exit Sub
    End If
    
    promptText = BuildBatchJsonPrompt(refTable, paraTexts)
    endpoint = gApiUrl & "/v1/chat/completions"
    requestJson = BuildChatCompletionJson_JSONMode(gModelName, promptText, gTemperature, gMaxTokens)
    
    On Error GoTo FailHandler
    Application.ScreenUpdating = False
    StatusBar = "Calling model with a " & gTimeoutSec & "s timeout ..."
    
    Dim waitStart As Single
    Dim waitedSec As Long
    Dim ok As Boolean
    Dim spinner As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, 30000, gTimeoutSec * 1000
    http.Open "POST", endpoint, True
    http.SetRequestHeader "Content-Type", "application/json"
    If Len(Trim$(gApiKey)) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & Trim$(gApiKey)
    End If
    
    http.Send requestJson

    waitStart = Timer

    Do
        DoEvents
    
        On Error Resume Next
        ok = http.WaitForResponse(1)
        If Err.Number <> 0 Then
            MsgBox "WaitForResponse failed:" & vbCrLf & _
                   "No.: " & Err.Number & vbCrLf & _
                   "Text: " & Err.Description, vbCritical
            On Error GoTo FailHandler
            GoTo CleanExit
        End If
        On Error GoTo FailHandler
    
        waitedSec = CLng(Timer - waitStart)
        spinner = Choose((waitedSec Mod 4) + 1, "|", "/", "-", "\")
        Application.StatusBar = "Calling model " & spinner & "  elapsed: " & waitedSec & " s"
    
        If ok Then Exit Do
    Loop

    If http.Status <> 200 Then
        MsgBox "HTTP error " & http.Status & vbCrLf & http.responseText, vbCritical
        GoTo CleanExit
    End If

    responseJson = http.responseText
    
    Dim finishReason As String

    finishReason = LCase$(Trim$(GetFinishReason(responseJson)))
    If finishReason = "length" Then
        MsgBox "Model output was truncated (finish_reason = length). Increase max_tokens and try again.", vbCritical
        GoTo CleanExit
    End If
    
    assistantText = ExtractAssistantContent(responseJson)
    assistantText = CleanupModelOutput(assistantText)
    
    If Trim$(assistantText) = "" Then
        MsgBox "Model returned empty assistant content.", vbCritical
        GoTo CleanExit
    End If

    If InStr(1, assistantText, """paragraphs""", vbTextCompare) = 0 Then
        MsgBox "Assistant content does not contain a paragraphs key." & vbCrLf & vbCrLf & _
               Left$(assistantText, 1200), vbCritical
        GoTo CleanExit
    End If
    
    Set rewrittenParas = ParseParagraphsFromJsonObject(assistantText)
    
    If rewrittenParas Is Nothing Then
        MsgBox "Could not parse JSON object with paragraphs array from the model answer.", vbCritical
        GoTo CleanExit
    End If
    
    Dim filteredParas As New Collection
    Dim k As Long
    Dim s As String

    For k = 1 To rewrittenParas.Count
        s = NormalizeParagraphText(CStr(rewrittenParas(k)))
        If IsSubstantiveParagraph(s) Then
            filteredParas.Add s
        End If
    Next k

    Set rewrittenParas = filteredParas
    
    If rewrittenParas.Count <> paraTexts.Count Then
      If gDebug Then
         MsgBox "Warning: Paragraph count mismatch." & vbCrLf & _
            "Sent: " & paraTexts.Count & vbCrLf & _
           "Returned: " & rewrittenParas.Count & vbCrLf & vbCrLf & _
           "Continuing with sequential matching.", vbExclamation
      End If
    End If
    
    oldTrackRevisions = targetDoc.TrackRevisions
    targetDoc.TrackRevisions = True
    
    failedParagraphs = ""
    
    Dim matchedPara As Collection
    Set matchedPara = MatchModelParasSequentially(paraTexts, rewrittenParas)

    For i = 1 To paraRanges.Count
        StatusBar = "Inserting reference signs in paragraph " & i & " of " & paraRanges.Count & "..."
    
        If CStr(matchedPara(i)) = "" Then
            If failedParagraphs <> "" Then failedParagraphs = failedParagraphs & ", "
            failedParagraphs = failedParagraphs & CStr(i)
        Else
            Set workRng = paraRanges(i).Duplicate
        
            If Len(workRng.Text) > 0 Then
                If Right$(workRng.Text, 1) = vbCr Then
                    workRng.End = workRng.End - 1
                End If
            End If
        
            okInsert = InsertReferenceSignsOnly( _
                workRng, _
                CStr(paraTexts(i)), _
                StripLeadingBracketNumber(FixNumberingTab(NormalizeParagraphText(CStr(matchedPara(i))))) _
            )
            
            If Not okInsert Then
                If failedParagraphs <> "" Then failedParagraphs = failedParagraphs & ", "
                failedParagraphs = failedParagraphs & CStr(i)
                MsgBox "Failure to rewrite: Original: " & paraTexts(i) & " Model: " & matchedPara(i)
            End If
        End If
    Next i
 
    
    targetDoc.TrackRevisions = oldTrackRevisions
    
    If failedParagraphs <> "" Then
        MsgBox "Done, but these paragraphs were skipped because the word sequence could not be aligned safely:" & vbCrLf & _
               failedParagraphs, vbExclamation
    End If

CleanExit:
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Exit Sub

FailHandler:
    On Error Resume Next
    If Not http Is Nothing Then http.Abort
    If Not targetDoc Is Nothing Then targetDoc.TrackRevisions = oldTrackRevisions
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    MsgBox "Request failed:" & vbCrLf & _
           "No.: " & Err.Number & vbCrLf & _
           "Text: " & Err.Description & vbCrLf, vbCritical
End Sub

Private Function InsertReferenceSignsOnly(ByVal targetRng As Range, ByVal originalText As String, ByVal modelText As String) As Boolean
    Dim oWordsPos As Collection
    Dim oWordsCmp As Collection
    Dim mWordsCmp As Collection
     
    Dim modelAnalysisText As String
    
    Dim iO As Long
    Dim iM As Long
    Dim insertAt As Long
    Dim delta As Long
    Dim refText As String
    Dim insRng As Range
    
    Dim canonO As String
    Dim canonM As String
    Dim combinedM As String
    Dim startM As Long
    Dim j As Long
    
    Dim dbgOriginalToken As String
    Dim dbgModelToken As String
    Dim dbgCombined As String
    
    modelAnalysisText = NormalizeAnalysisText(modelText)
    
    If gDebug Then
        MsgBox modelAnalysisText, vbCritical
    End If
    
    Set oWordsPos = ExtractWordsOnly(originalText)
    Set oWordsCmp = oWordsPos
    Set mWordsCmp = ExtractWordsOnly(modelAnalysisText)
    
    If mWordsCmp.Count = 0 Then
      If gDebug Then
        MsgBox "DEBUG 2 - No model words could be extracted." & vbCrLf & _
               "modelText: " & modelText, vbCritical
      End If

      InsertReferenceSignsOnly = False
      Exit Function
    End If
  
    
    iO = 1
    iM = 1
    delta = 0
    
    Do While iO <= oWordsCmp.Count And iM <= mWordsCmp.Count
        canonO = CanonicalWordForCompare(CStr(oWordsPos(iO)(0)))
        canonM = CanonicalWordForCompare(CStr(mWordsCmp(iM)(0)))
        
        dbgOriginalToken = CStr(oWordsPos(iO)(0))
        dbgModelToken = CStr(mWordsCmp(iM)(0))
        
        
        If canonO = canonM Then
            refText = CStr(mWordsCmp(iM)(3))
            
            If Len(refText) > 0 Then
                insertAt = CLng(oWordsPos(iO)(2)) + delta
                
                If Not AlreadyHasEquivalentRefsAfter(targetRng.Text, insertAt, refText) Then
                    Set insRng = targetRng.Duplicate
                    insRng.SetRange targetRng.Start + insertAt, targetRng.Start + insertAt
                    insRng.InsertAfter " " & refText
                    delta = delta + Len(refText) + 1
                End If
            End If
            
            iO = iO + 1
            iM = iM + 1
        
        Else
            
            If gDebug Then
              MsgBox "DEBUG 3 - Direct token mismatch, trying merge" & vbCrLf & _
                "iO = " & iO & ", iM = " & iM & vbCrLf & _
                "Original token = [" & dbgOriginalToken & "]" & vbCrLf & _
                "Original canonical = [" & canonO & "]" & vbCrLf & _
                "Model token = [" & dbgModelToken & "]" & vbCrLf & _
                "Model canonical = [" & canonM & "]" & vbCrLf & _
                "Original code points = " & CharCodes(dbgOriginalToken) & vbCrLf & _
                "Model code points = " & CharCodes(dbgModelToken) & vbCrLf & _
                "Original analysis = " & CharCodes(CStr(oWordsCmp(iO)(0))) & vbCrLf & _
                "Model analysis = " & CharCodes(CStr(mWordsCmp(iM)(0))) & vbCrLf, vbExclamation
            End If
            
            combinedM = canonM
            startM = iM
                
            dbgCombined = "[" & CStr(mWordsCmp(iM)(0)) & "]"
                
            For j = iM + 1 To mWordsCmp.Count
                If j > iM + 3 Then Exit For
                combinedM = combinedM & CanonicalWordForCompare(CStr(mWordsCmp(j)(0)))
                dbgCombined = dbgCombined & " + [" & CStr(mWordsCmp(j)(0)) & "]"
    
                If combinedM = canonO Then
                    refText = CollectRefsForMatchedWords(mWordsCmp, startM, j)
                    
                    If Len(refText) > 0 Then
                        insertAt = CLng(oWordsPos(iO)(2)) + delta
                        
                        If Not AlreadyHasEquivalentRefsAfter(targetRng.Text, insertAt, refText) Then
                            Set insRng = targetRng.Duplicate
                            insRng.SetRange targetRng.Start + insertAt, targetRng.Start + insertAt
                            insRng.InsertAfter " " & refText
                            delta = delta + Len(refText) + 1
                        End If
                    End If
                    
                    iO = iO + 1
                    iM = j + 1
                    GoTo NextLoop
                End If
                
                If Len(combinedM) >= Len(canonO) Then Exit For
            Next j
            
            If gDebug Then
              MsgBox "DEBUG 4 - Merge failed" & vbCrLf & _
                "Could not match original token after trying model-token merge." & vbCrLf & vbCrLf & _
                "iO = " & iO & ", iM = " & iM & vbCrLf & _
                "Original token = [" & CStr(oWordsCmp(iO)(0)) & "]" & vbCrLf & _
                "Original canonical = [" & canonO & "]" & vbCrLf & _
                "Tried model tokens = " & dbgCombined & vbCrLf & _
                "Final combined canonical = [" & combinedM & "]", vbCritical
            End If
            
            InsertReferenceSignsOnly = False
            Exit Function
        End If
        
NextLoop:
    Loop
   
    Rem If iO <= oWordsCmp.Count Or iM <= mWordsCmp.Count Then
    Rem    MsgBox "DEBUG 5 - Loop ended with leftover tokens" & vbCrLf & _
    rem        "iO = " & iO & " of " & oWordsCmp.Count & vbCrLf & _
    rem        "iM = " & iM & " of " & mWordsCmp.Count, vbCritical
    Rem    InsertReferenceSignsOnly = False
    Rem    Exit Function
    Rem End If
       
    InsertReferenceSignsOnly = True
End Function

Private Function TokenCollectionsMatchCount(ByVal a As Collection, ByVal b As Collection) As Boolean
    TokenCollectionsMatchCount = (a.Count = b.Count)
End Function


Private Function GetRefsAfterWord(ByVal s As String, ByVal wordEnd As Long) As String
    Dim p As Long
    Dim oneRef As String
    Dim refs As String

    p = wordEnd + 1

    Do
        Do While p <= Len(s) _
           And (Mid$(s, p, 1) = " " Or Mid$(s, p, 1) = vbTab)
            p = p + 1
        Loop

        oneRef = ReadParenthesizedGroup(s, p)
        If Len(oneRef) = 0 Then Exit Do

        refs = refs & oneRef
        p = p + Len(oneRef)
    Loop

    GetRefsAfterWord = refs
End Function


Private Function ExtractWordsOnly(ByVal s As String) As Collection
    Dim c As New Collection
    Dim i As Long
    Dim startPos As Long
    Dim token As String
    Dim item(3) As Variant
    Dim ch As String
    Dim depth As Long
    
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        
        If ch = "(" Then
            depth = 1
            i = i + 1
            
            Do While i <= Len(s) And depth > 0
                ch = Mid$(s, i, 1)
                
                If ch = "(" Then
                    depth = depth + 1
                ElseIf ch = ")" Then
                    depth = depth - 1
                End If
                
                i = i + 1
            Loop
            
        ElseIf IsWordChar(ch) Then
            startPos = i
            token = ""
            
            Do While i <= Len(s) And IsWordChar(Mid$(s, i, 1))
                token = token & Mid$(s, i, 1)
                i = i + 1
            Loop
            
            item(0) = token
            item(1) = startPos
            item(2) = i - 1
            item(3) = GetRefsAfterWord(s, item(2))
            c.Add item
            
        Else
            i = i + 1
        End If
    Loop
    
    Set ExtractWordsOnly = c
End Function

Private Function ReadParenthesizedGroup(ByVal s As String, ByVal startPos As Long) As String
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    
    If startPos < 1 Or startPos > Len(s) Then Exit Function
    If Mid$(s, startPos, 1) <> "(" Then Exit Function
    
    depth = 0
    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "(" Then
            depth = depth + 1
        ElseIf ch = ")" Then
            depth = depth - 1
            If depth = 0 Then
                ReadParenthesizedGroup = Mid$(s, startPos, i - startPos + 1)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function AlreadyHasEquivalentRefsAfter(ByVal s As String, ByVal insertPos As Long, ByVal refText As String) As Boolean
    Dim p As Long
    Dim actual As String
    
    ' Start where we will insert
    p = insertPos + 1
    
    ' Skip one existing space if present (because we always insert " " & refText)
    If p <= Len(s) Then
        If Mid$(s, p, 1) = " " Or Mid$(s, p, 1) = vbTab Then
            p = p + 1
        End If
    End If
    
    actual = ReadParenthesizedGroup(s, p)
    AlreadyHasEquivalentRefsAfter = (Len(actual) > 0 And actual = refText)
End Function

Private Function IsWhitespaceOnly(ByVal ch As String) As Boolean
    IsWhitespaceOnly = (ch = " " Or ch = vbTab)
End Function

Private Function IsWhitespaceChar(ByVal ch As String) As Boolean
    IsWhitespaceChar = ( _
        ch = " " _
        Or ch = vbTab _
        Or ch = vbCr _
        Or ch = vbLf _
    )
End Function

Private Function IsTokenPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", ")", "[", "]", "{", "}", _
             ".", ",", ";", ":", "!", "?", _
             """", "“", "”", "„", "«", "»", _
             "+", "=", "*", "&", "|", "\", "<", ">", _
             "@", "#", "$", "%", "^", "_", "~", "`"
            IsTokenPunctuation = True
        Case Else
            IsTokenPunctuation = False
    End Select
End Function

Private Function IsWordChar(ByVal ch As String) As Boolean
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

Private Function BuildBatchJsonPrompt(ByVal refTable As String, ByVal paraTexts As Collection) As String
    Dim s As String
    Dim i As Long
    
    s = ""
    s = s & "You are editing patent claims." & vbLf
    s = s & "Your task is to reproduce each paragraph exactly, preserving wording, numbering, punctuation, capitalization, and spacing as much as possible." & vbLf
    s = s & "Only insert reference signs in parentheses after the corresponding claim features." & vbLf
    s = s & "If the same term has multiple reference signs, add the first/lowest reference sign if the term is in singular, and add all reference signs as a comma-separated list inside parentehses if the term is in plural." & vbLf
    s = s & "The reference sign table may comprise further prompts for you to consider." & vbLf
    s = s & "Do not explain anything." & vbLf
    s = s & "Do not add commentary." & vbLf
    s = s & "Do not omit text." & vbLf
    s = s & "Do not rewrite or improve the claim language." & vbLf
    s = s & "Do not merge paragraphs." & vbLf
    s = s & "Do not split paragraphs." & vbLf
    s = s & "If a feature already has a reference sign, keep it and do not duplicate it." & vbLf
    s = s & "Return ONLY valid JSON." & vbLf
    s = s & "The JSON must have exactly one top-level object with one key named ""paragraphs""." & vbLf
    s = s & "The value of ""paragraphs"" must be an array of strings." & vbLf
    s = s & "The array must contain exactly " & CStr(paraTexts.Count) & " strings." & vbLf
    s = s & "Each array element must be a JSON string and nothing else." & vbLf
    s = s & "Each string must be the rewritten version of the corresponding input paragraph in the same order." & vbLf
    s = s & "Do not use markdown." & vbLf
    s = s & "Do not use code fences." & vbLf
    s = s & "Do not output any text before or after the JSON." & vbLf
    s = s & vbLf
    s = s & "Required output example:" & vbLf
    s = s & "{""paragraphs"":[""paragraph one..."",""paragraph two...""]}" & vbLf
    s = s & vbLf
    s = s & "Reference sign table:" & vbLf
    s = s & refTable & vbLf & vbLf
    s = s & "Paragraphs to rewrite:" & vbLf & vbLf
    
    For i = 1 To paraTexts.Count
        s = s & "[" & CStr(i) & "] " & paraTexts(i) & vbLf & vbLf
        Rem s = s & paraTexts(i) & vbLf & vbLf
    Next i
    
    BuildBatchJsonPrompt = s
End Function

Private Function BuildChatCompletionJson_JSONMode(ByVal modelName As String, ByVal promptText As String, ByVal temperature As Double, ByVal maxTokens As Long) As String
    Dim systemMsg As String
    Dim userMsg As String
    Dim json As String
    
    systemMsg = "You are a careful patent-editing assistant. Output only valid JSON with a top-level key named paragraphs."
    userMsg = promptText
    
    json = "{"
    json = json & """model"":""" & JsonEscape(modelName) & ""","
    json = json & """temperature"":" & FormatDotDouble(temperature) & ","
    json = json & """max_tokens"":" & CStr(maxTokens) & ","
    json = json & """response_format"":{""type"":""json_object""},"
    If (gThinking) Then
        json = json & """chat_template_kwargs"":{""enable_thinking"":true},"
    Else
        json = json & """chat_template_kwargs"":{""enable_thinking"":false},"
    End If
    json = json & """messages"":["
    json = json & "{""role"":""system"",""content"":""" & JsonEscape(systemMsg) & """},"
    json = json & "{""role"":""user"",""content"":""" & JsonEscape(userMsg) & """}"
    json = json & "]"
    json = json & "}"
    
    BuildChatCompletionJson_JSONMode = json
End Function

Private Function ShowLargeTextDialog(ByVal dialogTitle As String, ByVal dialogText As String, ByVal allowCancel As Boolean) As Boolean
    Dim f As frmPromptPreview
    
    Set f = New frmPromptPreview
    f.Caption = dialogTitle
    f.txtPreview.Text = dialogText
    f.Show vbModal
    
    ShowLargeTextDialog = Not f.Cancelled
    
    Unload f
    Set f = Nothing
End Function

Private Function GetReferenceListFromForm() As String
    Dim f As frmRefList
    
    Set f = New frmRefList
    f.txtRefList.Text = "vehicle" & vbTab & "10" & vbCrLf & _
                        "wheel" & vbTab & "2" & vbCrLf & _
                        "engine" & vbTab & "3" & vbCrLf & _
                        "surface" & vbTab & "100" & vbCrLf
    f.Show vbModal
    
    If f.Cancelled Then
        GetReferenceListFromForm = ""
    Else
        GetReferenceListFromForm = f.txtRefList.Text
    End If
    
    Unload f
    Set f = Nothing
End Function

Private Function ParseParagraphsFromJsonObject(ByVal jsonText As String) As Collection
    Dim pKey As Long
    Dim pColon As Long
    Dim arrStart As Long
    Dim arrEnd As Long
    Dim arrText As String
    
    jsonText = CleanupModelOutput(jsonText)
    
    pKey = InStr(1, jsonText, """paragraphs""", vbTextCompare)
    If pKey = 0 Then Exit Function
    
    pColon = InStr(pKey, jsonText, ":")
    If pColon = 0 Then Exit Function
    
    arrStart = InStr(pColon + 1, jsonText, "[")
    If arrStart = 0 Then Exit Function
    
    arrEnd = FindMatchingBracket(jsonText, arrStart)
    If arrEnd = 0 Then Exit Function
    
    arrText = Mid$(jsonText, arrStart + 1, arrEnd - arrStart - 1)
    Set ParseParagraphsFromJsonObject = ParseJsonStringArray(arrText)
End Function

Private Function ParseJsonStringArray(ByVal s As String) As Collection
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

Private Function FindMatchingBracket(ByVal s As String, ByVal openPos As Long) As Long
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

Private Function GetParagraphTextWithoutMark(ByVal rng As Range) As String
    Dim s As String
    s = rng.Text
    If Len(s) > 0 Then
        If Right$(s, 1) = vbCr Then
            s = Left$(s, Len(s) - 1)
        End If
    End If
    GetParagraphTextWithoutMark = s
End Function

Private Function IsSubstantiveParagraph(ByVal s As String) As Boolean
    Dim t As String
    t = Replace(s, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Trim$(t)
    IsSubstantiveParagraph = (t <> "")
End Function

Private Function NormalizeParagraphText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    
    Do While Len(s) > 0 And Right$(s, 1) = vbLf
        s = Left$(s, Len(s) - 1)
    Loop
    
    NormalizeParagraphText = s
End Function



Private Function ExtractAssistantContent(ByVal jsonText As String) As String
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
    
    pValue = InStr(pContent + 9, jsonText, """")
    If pValue = 0 Then Exit Function
    pValue = pValue + 1
    
    qValue = FindJsonStringEnd(jsonText, pValue)
    If qValue = 0 Then Exit Function
    
    ExtractAssistantContent = JsonUnescape(Mid$(jsonText, pValue, qValue - pValue))
End Function


Private Function FindJsonStringEnd(ByVal s As String, ByVal startPos As Long) As Long
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

Private Function JsonEscape(ByVal s As String) As String
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
                result = result & "\"""
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

Private Function JsonUnescape(ByVal s As String) As String
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
                    result = result & """"
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

Private Function IsHex4(ByVal s As String) As Boolean
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

Private Function CleanupModelOutput(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Trim$(s)
    
    If Left$(s, 3) = "```" Then
        s = StripCodeFences(s)
    End If
    
    CleanupModelOutput = Trim$(s)
End Function

Private Function StripCodeFences(ByVal s As String) As String
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

Private Function FixNumberingTab(ByVal s As String) As String
    If Len(s) >= 4 Then
        s = Replace(s, ". " & vbTab, "." & vbTab)
    End If
    
    If Len(s) >= 4 Then
        If Mid$(s, 1, 1) Like "#" _
           And Mid$(s, 2, 1) = "." _
           And Mid$(s, 3, 1) = " " _
           And Mid$(s, 4, 1) <> vbTab Then
            s = Left$(s, 2) & vbTab & Mid$(s, 4)
        
        ElseIf Len(s) >= 5 _
           And Mid$(s, 1, 1) Like "#" _
           And Mid$(s, 2, 1) Like "#" _
           And Mid$(s, 3, 1) = "." _
           And Mid$(s, 4, 1) = " " _
           And Mid$(s, 5, 1) <> vbTab Then
            s = Left$(s, 3) & vbTab & Mid$(s, 5)
        End If
    End If
    
    FixNumberingTab = s
End Function


Private Function StripLeadingBracketNumber(ByVal s As String) As String
    Dim p As Long
    Dim n As String
    
    s = LTrim$(s)
    
    If Left$(s, 1) <> "[" Then
        StripLeadingBracketNumber = s
        Exit Function
    End If
    
    p = InStr(2, s, "]")
    If p = 0 Then
        StripLeadingBracketNumber = s
        Exit Function
    End If
    
    n = Mid$(s, 2, p - 2)
    If n <> "" And IsAllDigitsText(n) Then
        StripLeadingBracketNumber = LTrim$(Mid$(s, p + 1))
    Else
        StripLeadingBracketNumber = s
    End If
End Function

Private Function IsAllDigitsText(ByVal s As String) As Boolean
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
Private Function CharCodes(ByVal s As String) As String
    Dim i As Long
    Dim result As String

    For i = 1 To Len(s)
        If Len(result) > 0 Then result = result & " "
        result = result & "U+" & Right$("0000" & Hex$(AscW(Mid$(s, i, 1))), 4)
    Next i

    CharCodes = result
End Function


Private Function NormalizeAnalysisText(ByVal s As String) As String
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

Private Function CanonicalWordForCompare(ByVal s As String) As String
    s = NormalizeAnalysisText(s)
    
    s = Replace(s, "-", "")
    s = Replace(s, "'", "")
    s = Replace(s, "/", "")
    
    CanonicalWordForCompare = LCase$(s)
End Function

Private Function CollectRefsForMatchedWords( _
    ByVal words As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long) As String

    Dim i As Long
    Dim result As String
    Dim oneRef As String

    For i = firstIndex To lastIndex
        oneRef = CStr(words(i)(3))

        If Len(oneRef) > 0 Then
            If result = "" Then
                result = oneRef
            ElseIf InStr(1, result, oneRef, vbBinaryCompare) = 0 Then
                result = result & oneRef
            End If
        End If
    Next i

    CollectRefsForMatchedWords = result
End Function

Sub ShowSelectedCharCodes()
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


Private Function GetFinishReason(ByVal jsonText As String) As String
    Dim pChoices As Long
    Dim pFinish As Long
    Dim pColon As Long
    Dim pQuote1 As Long
    Dim pQuote2 As Long
    
    pChoices = InStr(1, jsonText, """choices""", vbTextCompare)
    If pChoices = 0 Then Exit Function
    
    pFinish = InStr(pChoices, jsonText, """finish_reason""", vbTextCompare)
    If pFinish = 0 Then Exit Function
    
    pColon = InStr(pFinish, jsonText, ":")
    If pColon = 0 Then Exit Function
    
    pQuote1 = InStr(pColon + 1, jsonText, """")
    If pQuote1 = 0 Then Exit Function
    
    pQuote2 = FindJsonStringEnd(jsonText, pQuote1 + 1)
    If pQuote2 = 0 Then Exit Function
    
    GetFinishReason = Mid$(jsonText, pQuote1 + 1, pQuote2 - pQuote1 - 1)
End Function

Private Function MatchModelParasSequentially( _
    ByVal origParas As Collection, _
    ByVal modelParas As Collection) As Collection

    Const THRESHOLD As Double = 0.7
    Const LOOK_AHEAD As Long = 3

    Dim result As New Collection
    Dim iO As Long
    Dim iM As Long
    Dim probeM As Long
    Dim probeO As Long
    Dim bestM As Long
    Dim bestO As Long
    Dim bestScore As Double
    Dim score As Double

    iO = 1
    iM = 1

    Do While iO <= origParas.Count

        If iM > modelParas.Count Then
            result.Add ""
            iO = iO + 1
            GoTo NextOriginal
        End If

        ' First preference: find current original paragraph within
        ' the next LOOK_AHEAD model paragraphs.
        bestM = 0
        bestScore = -1#

        For probeM = iM To modelParas.Count
            If probeM > iM + LOOK_AHEAD Then Exit For

            score = ParagraphSimilarityScore( _
                CStr(origParas(iO)), _
                CStr(modelParas(probeM)) _
            )

            If score > bestScore Then
                bestScore = score
                bestM = probeM
            End If
        Next probeM

        If bestScore >= THRESHOLD Then
            result.Add CStr(modelParas(bestM))
            iO = iO + 1
            iM = bestM + 1
            GoTo NextOriginal
        End If

        ' No model match for current original paragraph. Check whether
        ' the next original paragraph matches current model paragraph.
        ' If so, the model likely omitted the current original paragraph.
        bestO = 0
        bestScore = -1#

        For probeO = iO + 1 To origParas.Count
            If probeO > iO + LOOK_AHEAD Then Exit For

            score = ParagraphSimilarityScore( _
                CStr(origParas(probeO)), _
                CStr(modelParas(iM)) _
            )

            If score > bestScore Then
                bestScore = score
                bestO = probeO
            End If
        Next probeO

        If bestScore >= THRESHOLD Then
            ' Model omitted original paragraphs iO through bestO - 1.
            ' Return unmatched entries for those original paragraphs;
            ' do not consume the current model paragraph.
            Do While iO < bestO
                result.Add ""
                iO = iO + 1
            Loop

            GoTo NextOriginal
        End If

        ' Neither sequence contains a credible near-term match.
        ' Do not force an unsafe association.
        result.Add ""
        iO = iO + 1

NextOriginal:
    Loop

    Set MatchModelParasSequentially = result
End Function

Private Function ParagraphSimilarityScore(ByVal s1 As String, ByVal s2 As String) As Double
    Dim w1 As Collection
    Dim w2 As Collection
    Dim i As Long
    Dim j As Long
    Dim matches As Long
    Dim denom As Long
    Dim c1 As String
    Dim c2 As String
    
    Set w1 = ExtractCanonicalWordsForParagraph(s1)
    Set w2 = ExtractCanonicalWordsForParagraph(s2)
    
    If w1.Count = 0 Or w2.Count = 0 Then
        ParagraphSimilarityScore = 0#
        Exit Function
    End If
    
    i = 1
    j = 1
    matches = 0
    
    Do While i <= w1.Count And j <= w2.Count
        c1 = CStr(w1(i))
        c2 = CStr(w2(j))
        
        If c1 = c2 Then
            matches = matches + 1
            i = i + 1
            j = j + 1
        Else
            If i < w1.Count Then
                If CStr(w1(i + 1)) = c2 Then
                    i = i + 1
                    GoTo NextScoreLoop
                End If
            End If
            
            If j < w2.Count Then
                If c1 = CStr(w2(j + 1)) Then
                    j = j + 1
                    GoTo NextScoreLoop
                End If
            End If
            
            i = i + 1
            j = j + 1
        End If
        
NextScoreLoop:
        
    Loop
    
    denom = w1.Count
    If w2.Count < denom Then denom = w2.Count
    
    If denom = 0 Then
        ParagraphSimilarityScore = 0#
    Else
        ParagraphSimilarityScore = matches / denom
    End If
End Function

Private Function ExtractCanonicalWordsForParagraph(ByVal s As String) As Collection
    Dim result As New Collection
    Dim words As Collection
    Dim i As Long
    Dim w As String
    
    s = StripLeadingBracketNumber(s)
    s = NormalizeParagraphText(s)
    s = NormalizeAnalysisText(s)
    
    Set words = ExtractWordsOnly(s)
    
    For i = 1 To words.Count
        w = CanonicalWordForCompare(CStr(words(i)(0)))
        If Len(w) > 0 Then result.Add w
    Next i
    
    Set ExtractCanonicalWordsForParagraph = result
End Function
