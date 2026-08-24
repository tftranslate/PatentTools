Attribute VB_Name = "modRefSigns"
Option Explicit

' Delimiters used in the hardcoded user message.
Private Const DELIM_REFS_BEGIN  As String = "===== BEGIN REFERENCE SIGN LIST ====="
Private Const DELIM_REFS_END    As String = "===== END REFERENCE SIGN LIST ====="
Private Const DELIM_PARA_BEGIN  As String = "===== BEGIN PARAGRAPHS "
Private Const DELIM_PARA_END    As String = "===== END PARAGRAPHS "
Private Const DELIM_SUFFIX      As String = " ====="


Public Sub Insert_Reference_Signs()
    Dim refTable As String
    Dim targetRange As Range
    Dim paraRanges As Collection
    Dim paraTexts As Collection
    Dim p As Paragraph
    Dim paraText As String
    Dim systemPrompt As String
    Dim userMessage As String
    Dim endpoint As String
    Dim requestJson As String
    Dim baseUrl As String
    Dim usedNativeRoute As Boolean
    Dim renderedPrompt As String
    Dim assistantText As String
    Dim finishReason As String
    Dim streamError As String
    Dim rewrittenParas As Collection
    Dim i As Long
    Dim useSelection As Boolean
    Dim workRng As Range
    
    Dim targetDoc As Document
    Dim oldTrackRevisions As Boolean
    Dim trackRevisionsSaved As Boolean
    Dim failedParagraphs As String
    Dim okInsert As Boolean
    
    Rem Populate the g... global variables
    LoadPatentToolsSettings

    ' Progress reporting is pointless if the status bar is hidden.
    Application.DisplayStatusBar = True
   
    refTable = GetPersistedReferenceList()
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
    
    systemPrompt = BuildRefSignSystemPrompt(paraTexts.Count)
    userMessage = BuildBatchUserMessage(refTable, paraTexts)
    baseUrl = NormalizeApiBaseUrl(gApiUrl)

    ' Route selection. Both backends are supported:
    '
    '   llama.cpp  -> /apply-template + /completion, the only documented way to
    '                 receive real "prompt_progress" during prefill.
    '   Ollama     -> /v1/chat/completions, which reports no progress; the
    '                 status bar then shows a calibrated estimate.
    '
    ' The native route is skipped when thinking is enabled: it constrains output
    ' with json_schema from the very first token, which leaves no room for a
    ' reasoning block, and /apply-template takes no chat_template_kwargs.
    usedNativeRoute = False

    If Not gThinking Then
        If PT_HasNativeProgressApi(baseUrl) Then
            If PT_ApplyTemplate(baseUrl, _
                    BuildMessagesArrayJson(systemPrompt, userMessage), _
                    gApiKey, gTimeoutSec, renderedPrompt, streamError) Then
                endpoint = baseUrl & "/completion"
                requestJson = BuildCompletionJson_JSONMode( _
                    renderedPrompt, gTemperature, gMaxTokens)
                usedNativeRoute = True
            End If
        End If
    End If

    If Not usedNativeRoute Then
        ' Fallback, and the normal path for Ollama. Any error from the probe or
        ' the template call is deliberately discarded: it is not a failure.
        streamError = ""
        endpoint = baseUrl & "/v1/chat/completions"
        requestJson = BuildChatCompletionJson_JSONMode( _
            gModelName, systemPrompt, userMessage, gTemperature, gMaxTokens)
    End If
    
    If gDebug Then
        If Not ShowLargeTextDialog( _
            "Prompt preview - system message and user message", _
            "----- ENDPOINT -----" & vbCrLf & endpoint & vbCrLf & _
            "(" & IIf(usedNativeRoute, _
                      "llama.cpp native route, real prompt progress", _
                      "OpenAI-compatible route, estimated prompt progress") & _
            ")" & vbCrLf & vbCrLf & _
            "----- SYSTEM MESSAGE -----" & vbCrLf & _
            ToDisplayText(systemPrompt) & vbCrLf & vbCrLf & _
            "----- USER MESSAGE -----" & vbCrLf & _
            ToDisplayText(userMessage), True) Then
            Exit Sub
        End If
    End If
    
    On Error GoTo FailHandler
    
    ' Streaming call. ScreenUpdating stays on: the status bar must repaint while
    ' the response arrives, and no document changes happen during the request.
    If usedNativeRoute Then
        Application.StatusBar = "Calling model with prompt-progress reporting ..."
    Else
        Application.StatusBar = "Calling model (timeout " & CStr(gTimeoutSec) & _
                                " s per idle period) ..."
    End If
    
    If Not StreamChatCompletion( _
            endpoint, requestJson, _
            assistantText, finishReason, streamError, _
            gTimeoutSec, gApiKey) Then
        MsgBox "The model call failed:" & vbCrLf & vbCrLf & streamError, vbCritical, "PatentTools"
        GoTo CleanExit
    End If
    
    finishReason = LCase$(Trim$(finishReason))
    If finishReason = "length" Then
        MsgBox "Model output was truncated (finish_reason = length). Increase max_tokens and try again.", vbCritical
        GoTo CleanExit
    End If
    
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
    trackRevisionsSaved = True
    targetDoc.TrackRevisions = True
    
    ' Screen updates are only suppressed while the document is being modified.
    Application.ScreenUpdating = False
    
    failedParagraphs = ""
    
    Dim matchedPara As Collection
    Set matchedPara = MatchModelParasSequentially(paraTexts, rewrittenParas)

    For i = 1 To paraRanges.Count
        Application.StatusBar = "Inserting reference signs in paragraph " & i & " of " & paraRanges.Count & "..."
    
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
    trackRevisionsSaved = False
    
    If failedParagraphs <> "" Then
        MsgBox "Done, but these paragraphs were skipped because the word sequence could not be aligned safely:" & vbCrLf & _
               failedParagraphs, vbExclamation
    End If

CleanExit:
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Exit Sub

FailHandler:
    Dim errNumber As Long
    Dim errText As String

    errNumber = Err.Number
    errText = Err.Description

    On Error Resume Next
    If trackRevisionsSaved Then
        If Not targetDoc Is Nothing Then targetDoc.TrackRevisions = oldTrackRevisions
    End If
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Request failed:" & vbCrLf & _
           "No.: " & CStr(errNumber) & vbCrLf & _
           "Text: " & errText & vbCrLf, vbCritical
End Sub
Private Function GetPersistedReferenceList() As String
    ' Read the persisted reference list from document custom properties.
    Dim refList As String
    On Error Resume Next
    refList = ActiveDocument.CustomDocumentProperties("PatentToolsRefList").Value
    On Error GoTo 0
    GetPersistedReferenceList = refList
End Function

Public Sub PatentTools_EditReferenceSigns(control As IRibbonControl)
    ' Callback for the "Edit reference signs" ribbon button.
    ' Opens the reference sign list dialog directly (without auto-prompting).
    Dim f As frmRefList

    Set f = New frmRefList
    f.Show vbModal

    Unload f
    Set f = Nothing
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


'-----------------------------------------------------------------------
' Prompt assembly
'
' System message: user-editable, held in gPromptInsert (settings dialog ->
'                 txtPromptInsert, factory text in modPatentToolsconfig).
' User message:   hardcoded here, carries only the data (reference sign list
'                 and the paragraphs), wrapped in explicit delimiters.
'-----------------------------------------------------------------------

' Returns the system message: the user-editable prompt with the paragraph-count
' placeholder resolved. If the user removed the placeholder, the count is appended
' so the model still receives the required array length.
Private Function BuildRefSignSystemPrompt(ByVal paraCount As Long) As String
    Dim s As String

    s = gPromptInsert
    If Trim$(s) = "" Then s = DEF_PromptInsert()

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)

    If InStr(1, s, PROMPT_COUNT_TOKEN, vbBinaryCompare) > 0 Then
        s = Replace(s, PROMPT_COUNT_TOKEN, CStr(paraCount))
    Else
        s = s & vbLf & "The array must contain exactly " & CStr(paraCount) & " strings."
    End If

    BuildRefSignSystemPrompt = s
End Function

' Returns the user message: data only, with explicit start/end delimiters around
' the reference sign list and around each paragraph to be reproduced.
Private Function BuildBatchUserMessage(ByVal refTable As String, ByVal paraTexts As Collection) As String
    Dim s As String
    Dim i As Long

    s = ""
    s = s & "The reference sign list is enclosed between the delimiter lines """ & DELIM_REFS_BEGIN & """ and """ & DELIM_REFS_END & "." & vbLf
    s = s & "Each paragraph you must reproduce is enclosed between its own delimiter lines """ & DELIM_PARA_BEGIN & "n" & DELIM_SUFFIX & """ and """ & DELIM_PARA_END & "n" & DELIM_SUFFIX & """, where n is the paragraph number." & vbLf
    s = s & "The delimiter lines are markup only. Never reproduce a delimiter line and never treat it as claim text." & vbLf
    s = s & "There are " & CStr(paraTexts.Count) & " paragraphs, numbered 1 to " & CStr(paraTexts.Count) & ", and the output array must follow that order." & vbLf
    s = s & vbLf

    s = s & DELIM_REFS_BEGIN & vbLf
    s = s & refTable & vbLf
    s = s & DELIM_REFS_END & vbLf & vbLf

    For i = 1 To paraTexts.Count
        s = s & DELIM_PARA_BEGIN & CStr(i) & DELIM_SUFFIX & vbLf
        s = s & paraTexts(i) & vbLf
        s = s & DELIM_PARA_END & CStr(i) & DELIM_SUFFIX & vbLf & vbLf
    Next i

    BuildBatchUserMessage = s
End Function

Private Function BuildChatCompletionJson_JSONMode( _
    ByVal modelName As String, _
    ByVal systemMsg As String, _
    ByVal userMsg As String, _
    ByVal temperature As Double, _
    ByVal maxTokens As Long _
) As String

    Dim json As String

    json = "{"
    json = json & """model"":""" & JsonEscape(modelName) & ""","
    json = json & """temperature"":" & FormatDotDouble(temperature) & ","
    json = json & """max_tokens"":" & CStr(maxTokens) & ","
    json = json & """response_format"":{""type"":""json_object""},"

    ' Streaming is part of the request, not something a transport layer may patch
    ' into the finished JSON afterwards.
    ' "return_progress" makes llama.cpp emit "prompt_progress" chunks during
    ' prefill; servers that do not know the flag simply ignore it.
    json = json & """stream"":true,"
    json = json & """return_progress"":true,"

    ' llama.cpp emits an SSE comment ping every N seconds while the stream is
    ' silent, which keeps the connection observable during long prompt
    ' processing - and, more importantly, makes the server flush the response
    ' headers immediately instead of at the first generated token. Without it,
    ' WinINet blocks inside HttpSendRequest for the whole prefill and Word looks
    ' frozen. Servers that do not know the field ignore it; a server that
    ' rejects it gets a reduced retry from modHttpStream.
    json = json & """sse_ping_interval"":1,"

    If gThinking Then
        json = json & """chat_template_kwargs"":{""enable_thinking"":true},"
    Else
        json = json & """chat_template_kwargs"":{""enable_thinking"":false},"
    End If

    json = json & """messages"":" & BuildMessagesArrayJson(systemMsg, userMsg)
    json = json & "}"
    
    BuildChatCompletionJson_JSONMode = json
End Function

' The bare messages array, shared by the chat-completions body and by
' /apply-template, so both routes are guaranteed to send the same conversation.
Private Function BuildMessagesArrayJson(ByVal systemMsg As String, _
                                        ByVal userMsg As String) As String
    Dim json As String

    json = "["
    json = json & "{""role"":""system"",""content"":""" & _
                  JsonEscape(systemMsg) & """},"
    json = json & "{""role"":""user"",""content"":""" & _
                  JsonEscape(userMsg) & """}"
    json = json & "]"

    BuildMessagesArrayJson = json
End Function

' Body for the native llama.cpp POST /completion, which is the only documented
' endpoint that emits "prompt_progress" during prefill. The prompt must already
' have been rendered through /apply-template.
'
' Differences from the chat body that matter:
'   * "n_predict" instead of "max_tokens".
'   * "json_schema" instead of "response_format": /completion has no
'     response_format, it constrains output through grammar or json_schema.
'     The schema is stricter than {"type":"json_object"} - it requires exactly
'     the paragraphs array of strings the parser expects.
'   * "cache_prompt" so an unchanged system prompt is not re-evaluated on the
'     next run; the cached share is then visible in the progress display.
Private Function BuildCompletionJson_JSONMode( _
    ByVal promptText As String, _
    ByVal temperature As Double, _
    ByVal maxTokens As Long _
) As String

    Dim json As String

    json = "{"
    json = json & """temperature"":" & FormatDotDouble(temperature) & ","
    json = json & """n_predict"":" & CStr(maxTokens) & ","
    json = json & """cache_prompt"":true,"
    json = json & """stream"":true,"
    json = json & """return_progress"":true,"
    json = json & """sse_ping_interval"":1,"
    json = json & """json_schema"":{""type"":""object"",""properties"":{" & _
                  """paragraphs"":{""type"":""array""," & _
                  """items"":{""type"":""string""}}}," & _
                  """required"":[""paragraphs""]},"

    ' Must stay last: the transport layer scopes its flag edits to the region
    ' before the payload member, and treats "prompt" like "messages".
    json = json & """prompt"":""" & JsonEscape(promptText) & """"
    json = json & "}"

    BuildCompletionJson_JSONMode = json
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
    
    ' Was: InStr(pContent + 9, jsonText, """) - an unterminated string literal.
    pValue = InStr(pContent + 9, jsonText, Chr$(34))
    If pValue = 0 Then Exit Function
    pValue = pValue + 1
    
    qValue = FindJsonStringEnd(jsonText, pValue)
    If qValue = 0 Then Exit Function
    
    ExtractAssistantContent = JsonUnescape(Mid$(jsonText, pValue, qValue - pValue))
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
    
	pQuote1 = InStr(pColon + 1, jsonText, Chr$(34))
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
