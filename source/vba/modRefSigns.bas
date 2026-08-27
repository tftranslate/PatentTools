Attribute VB_Name = "modRefSigns"
Option Explicit

' Delimiters used in the hardcoded user message.
Private Const DELIM_REFS_BEGIN  As String = "===== BEGIN REFERENCE SIGN LIST ====="
Private Const DELIM_REFS_END    As String = "===== END REFERENCE SIGN LIST ====="
Private Const DELIM_PARA_BEGIN  As String = "===== BEGIN PARAGRAPH NO. "
Private Const DELIM_PARA_END    As String = "===== END PARAGRAPH NO. "
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
    '                 Thinking mode is controlled per-request via chat_template_kwargs.
    '   Ollama     -> /v1/chat/completions, which reports no progress; the
    '                 status bar then shows a calibrated estimate.
    '
    ' Route selection logic:
    '   1. If gLlamaNative = True AND server supports /apply-template endpoint
    '      -> Use native route (with enable_thinking passed via chat_template_kwargs)
    '   2. Otherwise
    '      -> Fall back to OpenAI-compatible /v1/chat/completions
    usedNativeRoute = False
    
    ' Debug: Log route selection decision
    If gDebug Then
        Dim routeDebug As String
        routeDebug = "=== ROUTE SELECTION DEBUG ===" & vbCrLf & _
                     "gLlamaNative setting: " & IIf(gLlamaNative, "True", "False") & vbCrLf & _
                     "gThinking setting: " & IIf(gThinking, "True", "False") & vbCrLf
        
        If gLlamaNative Then
            Dim hasApi As Boolean
            hasApi = PT_HasNativeProgressApi(baseUrl)
            routeDebug = routeDebug & "Server supports /apply-template: " & IIf(hasApi, "Yes", "No") & vbCrLf
            
            If Not hasApi Then
                routeDebug = routeDebug & "Reason for OpenAI fallback: Server does not expose /apply-template endpoint" & vbCrLf
            End If
        Else
            routeDebug = routeDebug & "Reason for OpenAI fallback: gLlamaNative is False in settings" & vbCrLf
        End If
        
        Debug.Print routeDebug
    End If

    If gLlamaNative Then
        If PT_HasNativeProgressApi(baseUrl) Then
            If PT_ApplyTemplate(baseUrl, _
                    BuildMessagesArrayJson(systemPrompt, userMessage), _
                    gApiKey, gTimeoutSec, renderedPrompt, streamError, _
                    gThinking) Then  ' Pass gThinking to control template rendering
                endpoint = baseUrl & "/completion"
                requestJson = BuildCompletionJson_JSONMode( _
                    renderedPrompt, gTemperature, gMaxTokens)
                usedNativeRoute = True
            Else
                If gDebug Then Debug.Print "Route Selection: PT_ApplyTemplate failed, falling back to OpenAI-compatible route"
            End If
        Else
            If gDebug Then Debug.Print "Route Selection: Server does not support /apply-template, using OpenAI-compatible route"
        End If
    Else
        If gDebug Then Debug.Print "Route Selection: gLlamaNative=False, using OpenAI-compatible route"
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
		If Not ShowLargeTextDialog ("Raw json object sent to model", requestJson, True) Then
		   Exit Sub
		End if
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
	
	if gDebug Then
		If Not ShowLargeTextDialog ("Raw model output", assistantText, True) Then
		   Exit Sub
		End if
	end if
    
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
        ' Check if it's likely a truncation issue vs other parsing errors
        Dim jsonLen As Long
        jsonLen = Len(assistantText)
        
        If InStr(jsonLen - 20, assistantText, "}") = 0 Then
            ' JSON doesn't end with closing brace - likely truncated
            MsgBox "Model response was truncated. The JSON object is incomplete." & vbCrLf & vbCrLf & _
                   "Try increasing max_tokens (currently " & CStr(gMaxTokens) & _
                   ") or timeouts or check if your network connection is stable.", vbCritical, "Truncated Response"
        Else
            MsgBox "Could not parse JSON object with paragraphs array from the model answer." & vbCrLf & vbCrLf & _
                   "The response may contain invalid JSON characters or formatting." & vbCrLf & vbCrLf & _
                   "Check Debug mode to see the raw output.", vbCritical, "JSON Parse Error"
        End If
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
    
        If CStr(matchedPara(i)) = "" Or Len(Trim$(CStr(matchedPara(i)))) < 3 Then
            ' Match failed OR matched paragraph is essentially empty — check if source is also non-substantive
            If Not IsSubstantiveParagraph(paraTexts(i)) Or Len(Trim$(paraTexts(i))) < 3 Then
                ' Skip: both source and model are substantively empty — not a failure
            Else
                If gDebug Then
                    MsgBox "Paragraph " & i & ": Matching algorithm returned empty or negligible result." & vbCrLf & _
                           "Source paragraph length: " & Len(paraTexts(i)) & " characters" & vbCrLf & _
                           "This means similarity threshold (< 0.7) was not met for any model paragraph in the look-ahead window.", _
                           vbInformation, "Match Failure Debug"
                End If
                
                ' Source had content but model output is missing — record as failure
                If failedParagraphs <> "" Then failedParagraphs = failedParagraphs & ", "
                failedParagraphs = failedParagraphs & CStr(i)
            End If
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

Public Function Populate_Reference_Sign_Table() As String
    Dim workRange As Range
    Dim userMessage As String
    Dim systemPrompt As String
    Dim endpoint As String
    Dim requestJson As String
    Dim assistantText As String
    Dim finishReason As String
    Dim streamError As String
    Dim baseUrl As String
    Dim renderedPrompt As String
    Dim usedNativeRoute As Boolean
    
    ' 1. Load settings to initialize global variables (gApiUrl, gApiKey, etc.)
    LoadPatentToolsSettings
    
    ' 2. Determine the text to send as user message
    On Error Resume Next
    If Not Selection Is Nothing Then
        If Len(Trim$(Replace(Selection.Range.Text, vbCr, ""))) > 0 Then
            Set workRange = Selection.Range.Duplicate
        Else
            Set workRange = ActiveDocument.Content
        End If
    Else
        Set workRange = ActiveDocument.Content
    End If
    On Error GoTo 0
    
    userMessage = "Here is the patent description text." & vbCrLf & vbCrLf & "<PATENT_DESCRIPTION>" & vbCrLf & workRange.Text & vbCrLf & "</PATENT_DESCRIPTION>"
    
    ' 3. Build the request
    systemPrompt = gPromptPopulate
    If Trim$(systemPrompt) = "" Then systemPrompt = DEF_PromptPopulate()
    baseUrl = NormalizeApiBaseUrl(gApiUrl)

    ' Route selection. Both backends are supported:
    '
    '   llama.cpp  -> /apply-template + /completion, the only documented way to
    '                 receive real "prompt_progress" during prefill.
    '                 Thinking mode is controlled per-request via chat_template_kwargs.
    '   Ollama     -> /v1/chat/completions, which reports no progress; the
    '                 status bar then shows a calibrated estimate.
    '
    ' Route selection logic:
    '   1. If gLlamaNative = True AND server supports /apply-template endpoint
    '      -> Use native route (with enable_thinking passed via chat_template_kwargs)
    '   2. Otherwise
    '      -> Fall back to OpenAI-compatible /v1/chat/completions
    usedNativeRoute = False
    
    ' Debug: Log route selection decision for population process
    If gDebug Then
        Dim popRouteDebug As String
        popRouteDebug = "=== POPULATE ROUTE SELECTION DEBUG ===" & vbCrLf & _
                        "gLlamaNative setting: " & IIf(gLlamaNative, "True", "False") & vbCrLf & _
                        "gThinkPopulation setting: " & IIf(gThinkPopulation, "True", "False") & vbCrLf
        
        If gLlamaNative Then
            Dim hasApiPop As Boolean
            hasApiPop = PT_HasNativeProgressApi(baseUrl)
            popRouteDebug = popRouteDebug & "Server supports /apply-template: " & IIf(hasApiPop, "Yes", "No") & vbCrLf
            
            If Not hasApiPop Then
                popRouteDebug = popRouteDebug & "Reason for OpenAI fallback: Server does not expose /apply-template endpoint" & vbCrLf
            End If
        Else
            popRouteDebug = popRouteDebug & "Reason for OpenAI fallback: gLlamaNative is False in settings" & vbCrLf
        End If
        
        Debug.Print popRouteDebug
    End If

    If gLlamaNative Then
        If PT_HasNativeProgressApi(baseUrl) Then
            If PT_ApplyTemplate(baseUrl, _
                    BuildMessagesArrayJson(systemPrompt, userMessage), _
                    gApiKey, gTimeoutSecPopulate, renderedPrompt, streamError, _
                    gThinkPopulation) Then  ' Pass gThinkPopulation to control template rendering
                endpoint = baseUrl & "/completion"
                requestJson = BuildCompletionJson_PlaintextMode( _
                    renderedPrompt, gTempPopulate, gMaxTokens)
                usedNativeRoute = True
            Else
                If gDebug Then Debug.Print "Population Route Selection: PT_ApplyTemplate failed, falling back to OpenAI-compatible route"
            End If
        Else
            If gDebug Then Debug.Print "Population Route Selection: Server does not support /apply-template, using OpenAI-compatible route"
        End If
    Else
        If gDebug Then Debug.Print "Population Route Selection: gLlamaNative=False, using OpenAI-compatible route"
    End If

    If Not usedNativeRoute Then
        ' Fallback, and the normal path for Ollama. Any error from the probe or
        ' the template call is deliberately discarded: it is not a failure.
        streamError = ""
        endpoint = baseUrl & "/v1/chat/completions"
        requestJson = BuildChatCompletionJson_PlaintextMode( _
            gModelName, systemPrompt, userMessage, gTempPopulate, gMaxTokens)
    End If
        
    ' 4. Call the model
    If gDebug Then
        Dim debugInfo As String
        debugInfo = "=== POPULATE REQUEST JSON ===" & vbCrLf & vbCrLf
        debugInfo = debugInfo & requestJson
        ShowLargeTextDialog "Debug: Populate Request JSON", debugInfo, False
    End If
    
    If usedNativeRoute Then
        Application.StatusBar = "Calling model with prompt-progress reporting ..."
    Else
        Application.StatusBar = "Calling model (timeout " & CStr(gTimeoutSecPopulate) & _
                                " s per idle period) ..."
    End If
    
    If StreamChatCompletion( _
            endpoint, requestJson, _
            assistantText, finishReason, streamError, _
            gTimeoutSecPopulate, gApiKey) Then
        
        ' Debug: Show raw response FROM STREAM (before cleanup)
        If gDebug Then
            Call ShowLargeTextDialog("Raw Stream Response", "Character count: " & CStr(Len(assistantText)) & vbCrLf & vbCrLf & assistantText, True)
        End If
        
        ' Debug: Show raw model output (after cleanup but before normalization)
        If gDebug Then
            Dim cleanOutput As String
            cleanOutput = CleanupModelOutput(assistantText)
            assistantText = cleanOutput  ' Now use cleaned text for rest of processing
            
            If Not ShowLargeTextDialog("Raw Model Output (after cleanup)", cleanOutput, True) Then
                Populate_Reference_Sign_Table = "Debug: User cancelled"
                Exit Function
            End If
        Else
            assistantText = CleanupModelOutput(assistantText)
        End If
        
        ' Normalize table output: convert \t to TAB, sort by reference sign
        assistantText = NormalizePopulationTable(assistantText)
        
        ' Debug: Show normalized output (after normalization/sorting)
        If gDebug Then
            If Not ShowLargeTextDialog("Normalized Output (after sorting)", assistantText, True) Then
                Populate_Reference_Sign_Table = "Debug: User cancelled"
                Exit Function
            End If
        End If
        
        If Trim$(assistantText) = "" Then
            Populate_Reference_Sign_Table = "The model returned an empty response."
        Else
            Populate_Reference_Sign_Table = assistantText
        End If
    Else
        Populate_Reference_Sign_Table = "Model call failed: " & streamError
    End If
    
    Application.StatusBar = ""
End Function

'=======================================================================
' POPULATION TABLE NORMALIZATION
'=======================================================================

' Normalize population table output:
' - Convert \\t escape sequences to actual TAB characters
' - Sort table lines alphanumerically by reference sign (numeric-aware)
' - Only processes if at least 2 valid table lines detected; otherwise returns original unchanged
Private Function NormalizePopulationTable(ByVal assistantText As String) As String
    On Error GoTo Failed
   
    ' CRITICAL: Normalize line endings - convert vbLf to vbCrLf for consistent splitting
    ' Many models output just LF (vbLf) without CR, which Split(,, vbCrLf) won't detect
    Dim normalizedText As String
    normalizedText = Replace$(assistantText, vbLf, vbCrLf)
    normalizedText = Replace$(normalizedText, vbCr & vbCr, vbCrLf)  ' Handle edge case of double-CR
   
    Dim allLines() As String
    allLines = Split(normalizedText, vbCrLf)
   
    ' Debug: Log what we're receiving
    If gDebug Then
        Dim logMsg As String
        logMsg = "=== NORMALIZATION DEBUG ===" & vbCrLf & _
                 "Total lines in input (after LF→CRLF normalization): " & (UBound(allLines) + 1) & vbCrLf
        MsgBox logMsg
    End If
 
    If UBound(allLines) < 1 Then
        ' Less than 2 lines - nothing to process
        NormalizePopulationTable = assistantText
        Exit Function
    End If
    
    Dim tableLines As New Collection
    Dim observationLines As New Collection
    Dim i As Long
    Dim line As String
    Dim normalizedLine As String
    Dim isValid As Boolean
    Dim tableCount As Long, obsCount As Long  ' Debug counters
    
    ' Classify lines as table or observation
    For i = LBound(allLines) To UBound(allLines)
        line = allLines(i)
        
        ' Check if this is a table line
        isValid = IsTableLine(line, normalizedLine)
        
        If isValid Then
            tableLines.Add normalizedLine
            tableCount = tableCount + 1
        Else
            observationLines.Add line  ' Keep original for observations
            obsCount = obsCount + 1
        End If
    Next i
    
    ' Debug: Report classification results
    If gDebug Then
        Dim debugLog As String
        debugLog = "Classified " & CStr(tableCount) & " table lines and " & CStr(obsCount) & " observation lines" & vbCrLf & vbCrLf
        
        ' Check abort condition
        If tableCount < 2 Then
            debugLog = debugLog & "ABORT: Less than 2 table lines detected, returning original unchanged." & vbCrLf & vbCrLf
            debugLog = debugLog & "Sample of first input lines that were checked:" & vbCrLf
            Dim sampleLines As String
            Dim j As Long
            For j = LBound(allLines) To IIf(LBound(allLines) + 10 <= UBound(allLines), LBound(allLines) + 10, UBound(allLines))
                sampleLines = sampleLines & "  [" & allLines(j) & "]" & vbCrLf
            Next j
            debugLog = debugLog & sampleLines
        End If
        
        ' Show in dialog if debug enabled
        Call ShowLargeTextDialog("Normalization Classification Debug", debugLog, True)
    End If
    
    ' Abort silently: need at least 2 table lines to process
    If tableLines.Count < 2 Then
        NormalizePopulationTable = assistantText
        Exit Function
    End If
    
    ' Sort table lines with numeric-aware comparison
    Call SortTableLines(tableLines)
    
    ' Reconstruct output: sorted table + observations
    Dim finalOutput As String
    finalOutput = ""
    
    For i = 1 To tableLines.Count
        finalOutput = finalOutput & CStr(tableLines(i)) & vbCrLf
    Next i
    
    ' Add separator only if there are observations AND the first observation is not empty
    If observationLines.Count > 0 Then
        Dim firstObs As String
        firstObs = Trim$(CStr(observationLines(1)))
        
        ' Only add blank line separator if first observation has actual content
        If Len(firstObs) > 0 Then
            finalOutput = finalOutput & vbCrLf
        End If
    End If
    
    For i = 1 To observationLines.Count
        finalOutput = finalOutput & CStr(observationLines(i)) & vbCrLf
    Next i
    
    NormalizePopulationTable = finalOutput
    Exit Function
    
Failed:
    ' Silent failure - return original unchanged
    NormalizePopulationTable = assistantText
End Function

' Check if a line is a valid table line and normalize it
' Returns True/False via function, normalized line via ByRef parameter
Private Function IsTableLine(ByVal line As String, ByRef normalizedLine As String) As Boolean
    Dim prefix As String
    Dim rest As String
    Dim i As Long
    Dim ch As String
    Dim foundTab As Boolean
    Dim sepEnd As Long
	Dim checkLine as String
    
    ' Debug: Show what we're checking (first 80 chars)
    If gDebug And Len(line) > 0 Then
        checkLine = Left$(line, 80)
        ' Replace special chars for display
        checkLine = Replace$(checkLine, vbTab, "[TAB]")
        checkLine = Replace$(checkLine, vbCrLf, "[CR][LF]")
    End If
    
    ' Step 1: Extract alphanumeric prefix (1-5 characters)
    prefix = ""
    sepEnd = 0
    
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        
        If IsAlphanumeric(ch) And Len(prefix) < 5 Then
            prefix = prefix & ch
            sepEnd = i
        Else
            ' End of prefix region
            Exit For
        End If
    Next i
    
    ' Must have at least 1 character in prefix
    If Len(prefix) = 0 Or Len(prefix) > 5 Then
        IsTableLine = False
		
		if gDebug = 1 then
			MsgBox "Not a table line:" & vbCrLf & checkLine, vbCritical
		end if 
        Exit Function
    End If
    
    ' Step 2: Check for TAB or \t separator after the prefix
    rest = Mid$(line, sepEnd + 1)
    
    ' Look for actual TAB character (ASCII 9)
    Dim tabPos As Long
    tabPos = InStr(1, rest, Chr$(9))
    
    If tabPos > 0 And tabPos <= 3 Then
        ' Found TAB within first 3 chars of rest (allows some whitespace)
        Dim textPart As String
        textPart = Mid$(rest, tabPos + 1)
        normalizedLine = prefix & Chr$(9) & Trim$(textPart)
        IsTableLine = True
        Exit Function
    End If
    
    ' Look for escaped \t sequence (with optional whitespace)
    ' In VBA, backslash is not an escape character, so we use Chr$(92) for backslash
    rest = LTrim$(rest)
    
    ' Check for backslash + t as literal characters (two chars: \ and t)
    If Len(rest) >= 2 Then
        Dim firstTwo As String
        firstTwo = Left$(rest, 2)
        
        If firstTwo = Chr$(92) & "t" Then
            ' Exact \t with no leading space (after LTrim) - valid
            textPart = Mid$(rest, 3)
            normalizedLine = prefix & Chr$(9) & Trim$(textPart)
            IsTableLine = True
            Exit Function
        End If
    End If
    
	if gDebug = 1 then
	   MsgBox "Not a table line:" & vbCrLf & checkLine, vbCritical
	end if 
    
	' Not a table line
    IsTableLine = False
End Function

' Check if character is alphanumeric (digits 0-9 and letters A-Z, a-z)
Private Function IsAlphanumeric(ByVal ch As String) As Boolean
    ch = UCase$(ch)
    IsAlphanumeric = ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9"))
End Function

' Sort collection of table lines with numeric-aware comparison
Private Sub SortTableLines(ByRef lines As Collection)
    If lines.Count < 2 Then Exit Sub
    
    ' Convert to array for easier sorting
    Dim arr() As String
    ReDim arr(1 To lines.Count)
    Dim i As Long
    
    For i = 1 To lines.Count
        arr(i) = CStr(lines(i))
    Next i
    
    ' Bubble sort with numeric-aware comparison
    Dim j As Long
    Dim temp As String
    Dim swapped As Boolean
    
    Do
        swapped = False
        For i = 1 To UBound(arr) - 1
            If NaturalCompare(arr(i), arr(i + 1)) > 0 Then
                ' Swap
                temp = arr(i)
                arr(i) = arr(i + 1)
                arr(i + 1) = temp
                swapped = True
            End If
        Next i
    Loop While swapped
    
    ' Write back to collection
    Set lines = New Collection
    For i = 1 To UBound(arr)
        lines.Add arr(i)
    Next i
End Sub

' Numeric-aware string comparison for reference signs
' Returns: -1 if s1 < s2, 0 if equal, 1 if s1 > s2
Private Function NaturalCompare(ByVal s1 As String, ByVal s2 As String) As Integer
    Dim extractPrefix1 As String
    Dim extractPrefix2 As String
    Dim num1 As Long
    Dim num2 As Long
    Dim hasNum1 As Boolean
    Dim hasNum2 As Boolean
    
    ' Extract leading numeric portion from the reference sign (before TAB)
    extractPrefix1 = GetReferenceSign(s1)
    extractPrefix2 = GetReferenceSign(s2)
    
    ' Trim any whitespace that might have slipped through
    extractPrefix1 = Trim$(extractPrefix1)
    extractPrefix2 = Trim$(extractPrefix2)
    
    ' Extract ONLY the leading digits for numeric comparison
    Dim leadNum1 As String, leadNum2 As String
    Dim i As Long, ch As String
    
    leadNum1 = ""
    For i = 1 To Len(extractPrefix1)
        ch = Mid$(extractPrefix1, i, 1)
        If ch >= "0" And ch <= "9" Then
            leadNum1 = leadNum1 & ch
        Else
            Exit For
        End If
    Next i
    
    leadNum2 = ""
    For i = 1 To Len(extractPrefix2)
        ch = Mid$(extractPrefix2, i, 1)
        If ch >= "0" And ch <= "9" Then
            leadNum2 = leadNum2 & ch
        Else
            Exit For
        End If
    Next i
    
    ' Try to parse leading digits as numbers
    On Error Resume Next
    If Len(leadNum1) > 0 Then
        num1 = CLng(leadNum1)
        hasNum1 = (Err.Number = 0)
    Else
        hasNum1 = False
    End If
    Err.Clear
    
    If Len(leadNum2) > 0 Then
        num2 = CLng(leadNum2)
        hasNum2 = (Err.Number = 0)
    Else
        hasNum2 = False
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Both have leading numeric portions: compare as numbers
    If hasNum1 And hasNum2 Then
        If num1 < num2 Then
            NaturalCompare = -1
        ElseIf num1 > num2 Then
            NaturalCompare = 1
        Else
            NaturalCompare = 0
        End If
        Exit Function
    End If
    
    ' Only one has leading numeric: numeric comes first (e.g., "10" < "S130")
    If hasNum1 And Not hasNum2 Then
        NaturalCompare = -1
        Exit Function
    End If
    
    If Not hasNum1 And hasNum2 Then
        NaturalCompare = 1
        Exit Function
    End If
    
    ' Both non-numeric or both mixed: use string comparison
    Dim cmp As Long
    cmp = StrComp(extractPrefix1, extractPrefix2, vbTextCompare)
    
    If cmp < 0 Then
        NaturalCompare = -1
    ElseIf cmp > 0 Then
        NaturalCompare = 1
    Else
        NaturalCompare = 0
    End If
End Function

' Extract reference sign (text before TAB) from a table line
Private Function GetReferenceSign(ByVal tableLine As String) As String
    Dim tabPos As Long
    tabPos = InStr(1, tableLine, Chr$(9))
    
    If tabPos > 0 Then
        GetReferenceSign = Left$(tableLine, tabPos - 1)
    Else
        ' No TAB found - return entire line as fallback
        GetReferenceSign = tableLine
    End If
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
	Dim strDebug3 as String
	Dim strDebug4 as String
    
    modelAnalysisText = NormalizeAnalysisText(modelText) 
   
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
    
    ' Fallback skip counter: tolerate up to 2 total tokens skipped across the entire paragraph
    Dim totalSkipped As Long
    totalSkipped = 0
    
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
            
            ' --- Forward Skip Catch-Up: tolerate omission of up to 2 original tokens ---
            Dim skip As Long
            For skip = 1 To 2
                If iO + skip > oWordsPos.Count Then Exit For
                If CanonicalWordForCompare(CStr(oWordsPos(iO + skip)(0))) = canonM Then
                    ' Model omitted (skip) tokens starting at iO. Advance original past them.
                    iO = iO + skip + 1
                    iM = iM + 1
                    GoTo NextLoop
                End If
            Next skip
            
            If gDebug Then
              strDebug3 =  "DEBUG 3 - Direct token mismatch, trying merge" & vbCrLf & _
                "iO = " & iO & ", iM = " & iM & vbCrLf & _
                "Original token = [" & dbgOriginalToken & "]" & vbCrLf & _
                "Original canonical = [" & canonO & "]" & vbCrLf & _
                "Model token = [" & dbgModelToken & "]" & vbCrLf & _
                "Model canonical = [" & canonM & "]" & vbCrLf & _
                "Original code points = " & CharCodes(dbgOriginalToken) & vbCrLf & _
                "Model code points = " & CharCodes(dbgModelToken) & vbCrLf & _
                "Original analysis = " & CharCodes(CStr(oWordsCmp(iO)(0))) & vbCrLf & _
                "Model analysis = " & CharCodes(CStr(mWordsCmp(iM)(0))) & vbCrLf
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
              strDebug4 = "DEBUG 4 - Merge failed" & vbCrLf & _
                "Could not match original token after trying model-token merge." & vbCrLf & vbCrLf & _
                "iO = " & iO & ", iM = " & iM & vbCrLf & _
                "Original token = [" & CStr(oWordsCmp(iO)(0)) & "]" & vbCrLf & _
                "Original canonical = [" & canonO & "]" & vbCrLf & _
                "Tried model tokens = " & dbgCombined & vbCrLf & _
                "Final combined canonical = [" & combinedM & "]"
            End If

            ' Bi-directional skip ahead: tolerate a single misspelled or hallucinated token
            Dim skipAhead As Long
            For skipAhead = 1 To 2
                If iO + skipAhead > oWordsPos.Count Or iM + skipAhead > mWordsCmp.Count Then Exit For
                
                If CanonicalWordForCompare(CStr(oWordsPos(iO + skipAhead)(0))) = _
                   CanonicalWordForCompare(CStr(mWordsCmp(iM + skipAhead)(0))) Then
                    
                    ' Match found after skipping tokens on both sides.
                    iO = iO + skipAhead + 1
                    iM = iM + skipAhead + 1
                    GoTo NextLoop
                End If
            Next skipAhead
            
			If gDebug Then
			MsgBox strDebug3, vbCritical
			MsgBox strDebug4, vbCritical
			MsgBox "DEBUG 5 - Bidirectional lookahead failed", vbCritical
			End If

			' Final Fallback: Skip one token on each side if totalSkipped < 2
			' This handles cases where text corruption (e.g., literal Unicode escapes)
            ' prevents any other recovery strategy from working.
            If totalSkipped < 2 Then
                If gDebug Then
                    MsgBox "DEBUG 6 - Final fallback skip activated (total skipped so far = " & CStr(totalSkipped) & ")" & vbCrLf & _
                           "Original: [" & CStr(oWordsCmp(iO)(0)) & "]" & vbCrLf & _
                           "Model: [" & CStr(mWordsCmp(iM)(0)) & "]", vbInformation, "Fallback Skip"
                End If

                ' Skip one token on each side (counts as 2 total)
                iO = iO + 1
                iM = iM + 1
                totalSkipped = totalSkipped + 2

                If iO <= oWordsPos.Count And iM <= mWordsCmp.Count Then
                    GoTo NextLoop
                Else
                    ' Ran out of tokens after skip - now hard fail
                    If gDebug Then MsgBox "DEBUG 7 - Skipped to end of tokens, hard failing", vbCritical
                    InsertReferenceSignsOnly = False
                    Exit Function
                End If
            End If

            ' Total skip limit reached (2) - hard fail
            If gDebug Then
                MsgBox "DEBUG 8 - Skip limit (2) reached, cannot skip further. Hard failing.", vbCritical
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

    ' Universal override to suppress internal reasoning for models that use reasoning_effort (e.g., gpt-oss)
    'If Not gThinking Then
    '    json = json & """reasoning_effort"":""low"","
    'End If

    json = json & """messages"":" & BuildMessagesArrayJson(systemMsg, userMsg)
    json = json & "}"
    
    BuildChatCompletionJson_JSONMode = json
End Function

' Builds a chat completion request body specifically for the population process.
' Unlike the insertion version, it omits "response_format": "json_object" so that
' the model returns plain text (the markdown table of reference signs) instead of
' trying to format the output as a JSON structure.
Private Function BuildChatCompletionJson_PlaintextMode( _
    ByVal modelName As String, _
    ByVal systemMsg As String, _
    ByVal userMsg As String, _
    ByVal temperature As Double, _
    ByVal maxTokens As Long _
) As String

    Dim json As String

    json = "{"
    json = json & """model"":""" & JsonEscape(modelName) & """" & ","
    json = json & """temperature"": " & FormatDotDouble(temperature) & ","
    json = json & """max_tokens"": " & CStr(maxTokens) & ","

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

    If gThinkPopulation Then
        json = json & """chat_template_kwargs"":{""enable_thinking"":true},"
    Else
        json = json & """chat_template_kwargs"":{""enable_thinking"":false},"
    End If

    ' Universal override to suppress internal reasoning for models that use reasoning_effort (e.g., gpt-oss)
    If Not gThinkPopulation Then
        json = json & """reasoning_effort"":""low"","
    End If

    json = json & """messages"": " & BuildMessagesArrayJson(systemMsg, userMsg)
    json = json & "}"
    
    BuildChatCompletionJson_PlaintextMode = json
End Function

' The bare messages array, shared by the chat-completions body and by
' /apply-template, so both routes are guaranteed to send the same conversation.
Private Function BuildMessagesArrayJson(ByVal systemMsg As String, _
                                        ByVal userMsg As String) As String
    Dim json As String
    
    ' Normalize line endings: CRLF → LF for all model input
    ' Reduces transmitted data and ensures consistent token counting
    systemMsg = Replace$(systemMsg, vbCrLf, vbLf)
    userMsg = Replace$(userMsg, vbCrLf, vbLf)

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

' Body for the native llama.cpp POST /completion endpoint used when
' populating the reference-sign table. This deliberately has no
' json_schema because the population prompt requires plain-text output.
Private Function BuildCompletionJson_PlaintextMode( _
    ByVal promptText As String, _
    ByVal temperature As Double, _
    ByVal maxTokens As Long) As String

    Dim json As String

    json = "{"
    json = json & """temperature"":" & FormatDotDouble(temperature) & ","
    json = json & """n_predict"":" & CStr(maxTokens) & ","
    json = json & """cache_prompt"":true,"
    json = json & """stream"":true,"
    json = json & """return_progress"":true,"
    json = json & """sse_ping_interval"":1,"
    json = json & """prompt"":""" & JsonEscape(promptText) & """"
    json = json & "}"

    BuildCompletionJson_PlaintextMode = json
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
    Dim topLevelEnd As Long
	
	if gDebug Then
	endif
	
    jsonText = CleanupModelOutput(jsonText)
    
    pKey = InStr(1, jsonText, """paragraphs""", vbTextCompare)
    If pKey = 0 Then Exit Function
    
    pColon = InStr(pKey, jsonText, ":")
    If pColon = 0 Then Exit Function
    
    arrStart = InStr(pColon + 1, jsonText, "[")
    If arrStart = 0 Then Exit Function
    
    ' Check if the JSON response is complete: look for closing brace of top-level object
    topLevelEnd = FindMatchingBracket(jsonText, InStr(1, jsonText, "{"))
    
    If topLevelEnd = 0 Then
        ' Top-level object not properly closed - likely truncated response
        Exit Function
    End If
    
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
