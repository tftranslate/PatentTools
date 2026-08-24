Attribute VB_Name = "modPatentToolsconfig"
Option Explicit

'=======================================================================
' DECLARATIONS SECTION
' All module-level declarations (Public/Private/Const) MUST appear here,
' above the first procedure. VBA compiler rejects constants below procedures.
'=======================================================================

Public gApiUrl As String
Public gApiKey As String
Public gModelName As String
Public gTemperature As Double
Public gTimeoutSec As Long
Public gMaxTokens As Long
Public gThinking As Boolean
Public gDebug As Boolean
Public gPromptInsert As String

Private Const APP_NAME As String = "PatentTools"
Private Const SECTION_NAME As String = "Settings"

' Tool release version.
Public Const TOOL_VERSION As String = "0.2.0-beta"

' Single source of truth: factory-default values for all globally persisted settings.
Private Const DEF_ApiUrl      As String    = "http://localhost:11434"
Private Const DEF_ApiKey      As String    = ""
Private Const DEF_ModelName   As String    = "gemma4:12b"
Private Const DEF_Temperature As Double    = 0.2
Private Const DEF_TimeoutSec  As Long      = 120
Private Const DEF_MaxTokens   As Long      = 32768
Private Const DEF_Thinking    As Boolean   = False
Private Const DEF_Debug       As Boolean   = False
' DEF_PromptInsert is not a Const: the factory system prompt is a multi-line text and
' exceeds what a single Const statement can hold within VBA's line-continuation limit.
' Use DEF_PromptInsert() wherever the other DEF_* constants are used.

' Placeholder inside the system prompt that modRefSigns replaces with the actual number
' of paragraphs sent in the request. Keep it in the prompt text when editing the prompt.
Public Const PROMPT_COUNT_TOKEN As String = "{PARAGRAPH_COUNT}"

'=======================================================================
' FACTORY DEFAULT SYSTEM PROMPT
'=======================================================================

' Factory default for gPromptInsert: the system message sent to the model by
' modRefSigns.Insert_Reference_Signs. Editable by the user in the settings dialog
' (txtPromptInsert) and persisted machine-wide like every other setting.
Public Function DEF_PromptInsert() As String
    Dim s As String

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
    s = s & "The array must contain exactly " & PROMPT_COUNT_TOKEN & " strings." & vbLf
    s = s & "Each array element must be a JSON string and nothing else." & vbLf
    s = s & "Each string must be the rewritten version of the corresponding input paragraph in the same order." & vbLf
    s = s & "Do not use markdown." & vbLf
    s = s & "Do not use code fences." & vbLf
    s = s & "Do not output any text before or after the JSON." & vbLf
    s = s & vbLf
    s = s & "Required output example:" & vbLf
    s = s & "{""paragraphs"":[""paragraph one..."",""paragraph two...""]}"

    DEF_PromptInsert = s
End Function

'=======================================================================
' HELPERS
'=======================================================================

' Multi-line settings are stored as a single registry line: real line breaks are
' encoded, so the value survives GetSetting/SaveSetting round-trips unchanged.
Public Function EncodeMultiLineSetting(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")

    EncodeMultiLineSetting = s
End Function

Public Function DecodeMultiLineSetting(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim nxt As String
    Dim result As String

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = "\" And i < Len(s) Then
            nxt = Mid$(s, i + 1, 1)

            Select Case nxt
                Case "n"
                    result = result & vbLf
                    i = i + 2
                Case "t"
                    result = result & vbTab
                    i = i + 2
                Case "\"
                    result = result & "\"
                    i = i + 2
                Case Else
                    result = result & ch
                    i = i + 1
            End Select
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    DecodeMultiLineSetting = result
End Function

' Line breaks are normalized to vbLf internally; MSForms text boxes need vbCrLf.
Public Function ToDisplayText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)

    ToDisplayText = Replace(s, vbLf, vbCrLf)
End Function

Public Function FromDisplayText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)

    FromDisplayText = s
End Function

Public Function ParseDotDouble(ByVal s As String) As Double
    Dim localDecimal As String

    s = Trim$(s)
    localDecimal = Mid$(CStr(1.1), 2, 1)

    ParseDotDouble = CDbl(Replace(s, ".", localDecimal))
End Function

Public Function FormatDotDouble(ByVal dblValue As Double) As String
    Dim localDecimal As String

    localDecimal = Mid$(CStr(1.1), 2, 1)

    FormatDotDouble = Replace(CStr(dblValue), localDecimal, ".")
End Function

'=======================================================================
' PERSISTENCE
'=======================================================================

Public Sub LoadPatentToolsSettings()
    Dim s As String

    s = GetSetting(APP_NAME, SECTION_NAME, "ApiUrl", "")
    gApiUrl     = IIf(Len(s) > 0, s, DEF_ApiUrl)

    s = GetSetting(APP_NAME, SECTION_NAME, "ApiKey", "")
    gApiKey     = IIf(Len(s) > 0, s, DEF_ApiKey)

    s = GetSetting(APP_NAME, SECTION_NAME, "ModelName", "")
    gModelName  = IIf(Len(s) > 0, s, DEF_ModelName)

    s = GetSetting(APP_NAME, SECTION_NAME, "Temperature", "")
    gTemperature = IIf(Len(s) > 0, ParseDotDouble(s), DEF_Temperature)

    s = GetSetting(APP_NAME, SECTION_NAME, "TimeoutSec", "")
    gTimeoutSec = IIf(Len(s) > 0, CLng(s), DEF_TimeoutSec)

    s = GetSetting(APP_NAME, SECTION_NAME, "MaxTokens", "")
    gMaxTokens  = IIf(Len(s) > 0, CLng(s), DEF_MaxTokens)

    s = GetSetting(APP_NAME, SECTION_NAME, "Thinking", "")
    gThinking   = IIf(Len(s) > 0, CBool(Val(s)), DEF_Thinking)

    s = GetSetting(APP_NAME, SECTION_NAME, "Debug", "")
    gDebug      = IIf(Len(s) > 0, CBool(Val(s)), DEF_Debug)

    s = GetSetting(APP_NAME, SECTION_NAME, "PromptInsert", "")
    If Len(s) > 0 Then
        gPromptInsert = DecodeMultiLineSetting(s)
    Else
        gPromptInsert = DEF_PromptInsert()
    End If

    ' If the registry is completely empty (first run), write defaults so future loads persist.
    If GetSetting(APP_NAME, SECTION_NAME, "ApiUrl", "__NULL__") = "__NULL__" Then
        SavePatentToolsSettings
    End If
End Sub

' Resets all globally persisted settings in the Windows registry to their factory defaults.
Public Sub ResetSettingsToDefaults()
    gApiUrl     = DEF_ApiUrl
    gApiKey     = DEF_ApiKey
    gModelName  = DEF_ModelName
    gTemperature = DEF_Temperature
    gTimeoutSec = DEF_TimeoutSec
    gMaxTokens  = DEF_MaxTokens
    gThinking   = DEF_Thinking
    gDebug      = DEF_Debug
    gPromptInsert = DEF_PromptInsert()

    SavePatentToolsSettings
End Sub

Public Sub SavePatentToolsSettings()
    SaveSetting APP_NAME, SECTION_NAME, "ApiUrl", gApiUrl
    SaveSetting APP_NAME, SECTION_NAME, "ApiKey", gApiKey
    SaveSetting APP_NAME, SECTION_NAME, "ModelName", gModelName
    SaveSetting APP_NAME, SECTION_NAME, "Temperature", FormatDotDouble(gTemperature)
    SaveSetting APP_NAME, SECTION_NAME, "TimeoutSec", CStr(gTimeoutSec)
    SaveSetting APP_NAME, SECTION_NAME, "MaxTokens", CStr(gMaxTokens)
    SaveSetting APP_NAME, SECTION_NAME, "Thinking", IIf(gThinking, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "Debug", IIf(gDebug, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "PromptInsert", EncodeMultiLineSetting(gPromptInsert)
End Sub

Public Function NormalizeApiBaseUrl(ByVal s As String) As String
    s = Trim$(s)
    
    Do While Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    
    If LCase$(Right$(s, 20)) = "/v1/chat/completions" Then
        s = Left$(s, Len(s) - 20)
    ElseIf LCase$(Right$(s, 3)) = "/v1" Then
        s = Left$(s, Len(s) - 3)
    End If
    
    Do While Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    
    NormalizeApiBaseUrl = s
End Function

Public Sub Patent_Tools_Settings()
    frmPatentToolsSettings.Show
End Sub

Public Function FetchModelList(ByVal rawApiUrl As String, ByRef modelNames As Collection, ByRef errorText As String) As Boolean
    Dim base As String
    Dim url As String
    Dim http As Object
    Dim body As String
    Dim lowerBody As String
    Dim recvMs As Long
    Dim pData As Long
    Dim arrStart As Long
    Dim arrEnd As Long
    Dim scanSrc As String
    Dim i As Long
    Dim pId As Long
    Dim nextPos As Long
    Dim ch As String
    Dim colonFound As Boolean
    Dim openQuote As Long
    Dim qValEnd As Long
    Dim valStr As String
    Dim k As Long
    Dim isDup As Boolean
    
    ' GET {base}/v1/models and collect the "id" value of every entry in the model list.
    ' Reuses FindMatchingBracket / FindJsonStringEnd / JsonUnescape from this module; the scan
    ' below is deliberately heuristic, same pragmatic style as the other JSON handling here.
    
    Set modelNames = New Collection
    errorText = ""
    
    base = NormalizeApiBaseUrl(Trim$(rawApiUrl))
    If LCase$(Left$(base, 5)) <> "http:" And LCase$(Left$(base, 6)) <> "https:" Then
        errorText = "Enter a valid http(s) base URL first."
        Exit Function
    End If
    
    url = base & "/v1/models"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    recvMs = gTimeoutSec * 1000
    If recvMs <= 0 Then recvMs = 30000     ' settings not loaded yet: never wait forever
    If recvMs > 30000 Then recvMs = 30000   ' a model list should arrive quickly; keep the dialog responsive
    
    http.SetTimeouts 5000, 8000, 5000, recvMs
    
    On Error Resume Next
    http.Open "GET", url, False
    If Err.Number <> 0 Then
        errorText = "Could not reach the server at this URL."
        Exit Function
    End If
    If Len(Trim$(gApiKey)) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & Trim$(gApiKey)
    End If
    http.Send
    If Err.Number <> 0 Then
        errorText = "The request for the model list failed."
        Exit Function
    End If
    On Error GoTo 0
    
    If CInt(http.Status) <> 200 Then
        If http.Status = 401 Or http.Status = 403 Then
            errorText = "Access denied (HTTP " & CStr(http.Status) & "). Check URL and API key."
        ElseIf http.Status = 404 Then
            errorText = "This server has no /v1/models endpoint (HTTP 404)."
        Else
            errorText = "Server error (HTTP " & CStr(http.Status) & ")."
        End If
        Exit Function
    End If
    
    body = http.responseText
    lowerBody = LCase$(body)
    
    ' Primary shape: {"data": [ ... ]}; some servers answer with a bare top-level array.
    pData = InStr(1, lowerBody, """data""")
    arrStart = 0
    If pData > 0 Then
        arrStart = InStr(pData + 5, body, "[")
    Else
        If Left$(Trim$(body), 1) = "[" Then
            arrStart = InStr(1, body, "[")
        End If
    End If
    If arrStart = 0 Then
        errorText = "The server response contains no model list."
        Exit Function
    End If
    
    arrEnd = FindMatchingBracket(body, arrStart)
    If arrEnd < arrStart + 1 Then
        errorText = "Malformed model list in the server response."
        Exit Function
    End If
    
    scanSrc = Mid$(body, arrStart + 1, arrEnd - arrStart - 1)
    
    ' Structural JSON keys are the only unescaped occurrences of literal quoted text like "id"
    ' (quotes inside string values arrive escaped), so scanning for key-then-value pairs here is safe.
    i = 1
    Do While i <= Len(scanSrc)
        pId = InStr(i, scanSrc, """id""", vbBinaryCompare)
        If pId = 0 Then Exit Do
        
        colonFound = False
        openQuote = 0
        nextPos = pId + 4   ' first character after the closing quote of the "id" key
        Do While nextPos <= Len(scanSrc) And (Mid$(scanSrc, nextPos, 1) = " ")
            nextPos = nextPos + 1
        Loop
        If nextPos > Len(scanSrc) Then Exit Do   ' malformed: no colon after the key; nothing more to scan
        
        ch = Mid$(scanSrc, nextPos, 1)
        If ch <> ":" Then
            i = nextPos    ' not a key we understand (e.g. "identity"); skip this occurrence
        Else
            Do While nextPos < Len(scanSrc) And (Mid$(scanSrc, nextPos + 1, 1) = " ")
                nextPos = nextPos + 1
            Loop
            openQuote = nextPos + 1
            If InStr(1, Mid$(scanSrc, openQuote), Chr$(34)) = 0 Or Left$(Mid$(scanSrc, openQuote), 1) <> Chr$(34) Then
                i = pId + 4     ' value is not a JSON string (unexpected); skip this occurrence
            Else
                qValEnd = FindJsonStringEnd(scanSrc, openQuote + 1)
                If qValEnd < openQuote + 1 Then Exit Do   ' unterminated string: stop, let zero-found handle it
                valStr = JsonUnescape(Mid$(scanSrc, openQuote + 1, qValEnd - openQuote - 1))
                If Len(Trim$(valStr)) > 0 Then
                    isDup = False
                    For k = 1 To modelNames.Count
                        If LCase$(modelNames(k)) = LCase$(valStr) Then isDup = True
                    Next k
                    If Not isDup Then modelNames.Add valStr
                End If
                i = qValEnd + 1
            End If
        End If
    Loop
    
    If modelNames.Count = 0 Then
        errorText = "The server reported no models."
        Exit Function
    End If
    FetchModelList = True
    
    ' success: caller selects the first item and displays the green status line.
End Function



