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
Public gTempPopulate As Double
Public gTimeoutSecPopulate As Long
Public gThinkPopulation As Boolean
Public gDebug As Boolean
Public gPromptInsert   As String
Public gPromptPopulate As String

Private Const APP_NAME As String = "PatentTools"
Private Const SECTION_NAME As String = "Settings"

' Tool release version.
Public Const TOOL_VERSION As String = "0.2.0"

' Single source of truth: factory-default values for all globally persisted settings.
Private Const DEF_ApiUrl      As String    = "http://localhost:11434"
Private Const DEF_ApiKey      As String    = ""
Private Const DEF_ModelName   As String    = "gemma4:12b"
Private Const DEF_Temperature  As Double     = 0.2
Private Const DEF_TimeoutSec   As Long       = 300  ' Increased for insertion (was 120)
Private Const DEF_MaxTokens    As Long       = 32768
Private Const DEF_Thinking     As Boolean    = False
Private Const DEF_TempPopulate As Double     = 0.6
Private Const DEF_TimeoutSecPopulate As Long = 900  ' Increased for population (was 600)
Private Const DEF_ThinkPopulation As Boolean = True
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

' Factory default system prompt for the population process: scan a patent description
' for reference-sign usage and compile a table of designated elements (content of
' PopulationPrompt.md). Mirrors DEF_PromptInsert in mechanism; only the text differs.
Public Function DEF_PromptPopulate	() As String
    Dim s As String

    s = s & "# Task" & vbLf
    s = s & "The task is a data extraction task." & vbLf
    s = s & "Scan the patent description for reference signs and the elements designated by them." & vbLf
    s = s & "Compile a comprehensive table of all reference signs that occur in the patent description with elements designated thereby." & vbLf
    s = s & "The table will be used as a guidance to a patent paralegal tasked with inserting reference signs in parentheses into the claims under European patent practice." & vbLf
    s = s & vbLf
    s = s & "# Definitions" & vbLf
    s = s & "A ""designated element"" may be a noun, a compound noun, a combination of qualifier + noun/compound noun, or in some cases also a verb in -ing form. It is a structural element, functional element or a method step or decision block that is designed in the figures by the reference sign." & vbLf
    s = s & "A ""reference sign"" is usually a one to four digit number. Sometimes the reference sign may also comprise a letter, as in ""71a"" or ""S130"", or may only be uppercase letters, like in ""SW"" or ""C"". The reference sign may relate to any structural element, functional element or method step, process step or flowchart step." & vbLf
    s = s & vbLf
    s = s & "# Strategy for resolving issues" & vbLf
    s = s & "Resolve ambiguities and other issues by giving further observations below the table such that the paralegal does not struggle to pick the correct reference number for a given element." & vbLf
    s = s & vbLf
    s = s & "## Multiple reference signs for same element" & vbLf
    s = s & "Check if you overlooked a qualifier that is actually part of the designated element. For example, if ""device"" has signs 10 and 20, maybe ""control device 10"" and ""interface device 20"" or ""first device 10"" and ""second device 20"" is the solution. If so, adapt the table." & vbLf
    s = s & "Check if the reference sign depends on embodiment/context. If so give an observation identifying which of the multiple reference numbers should be used for the designated element in which claim and/or in which context." & vbLf
    s = s & vbLf
    s = s & "## Multiple elements having same reference sign" & vbLf
    s = s & "Determine what terminology is used in the claims." & vbLf
    s = s & "The designated element that is also in the claims is the designated element; any other designated elements are merely further observations." & vbLf
    s = s & vbLf
    s = s & "## Collective element versus individual element" & vbLf
    s = s & "Sometimes an element, such as a collective or generic element has one reference sign, and components or specific embodiments of the collective element that share portions of the name have another sign. Example: ""bin 1"" versus ""dust bin 101"" and ""rubbish bin 102"". Alert the user to such cases with a further observation." & vbLf
    s = s & vbLf
    s = s & "## Not an ambiguity" & vbLf
    s = s & "If an element is sometimes used with and sometimes without reference sign, this is not an ambiguity issue and there is nothing to observe. It is normal because when you work the claims do not yet have reference signs." & vbLf
    s = s & vbLf
    s = s & "# Output description" & vbLf
    s = s & "Do not output anything in front of the table." & vbLf
    s = s & "Output the table only." & vbLf
    s = s & "After the table, you may output fan empty line and, in starting from the next line after the empty line, further observations." & vbLf
    s = s & vbLf
    s = s & "## Table" & vbLf
    s = s & "The table must contain all reference signs that occur in the description. I.e., we want all reference signs for all structural elements, for all functional elements, for all flowchart steps and for all process steps and for all method steps designated with a reference sign anywhere in the description." & vbLf
    s = s & vbLf
    s = s & "Do not spend time sorting the table." & vbLf
    s = s & vbLf
	s = s & "Do not output a markdown table but just a plain text table formatted as follows:" & vbLf
    s = s & "Each table line starts with the reference sign, then a tabulator character, then the designated element." & vbLf
    s = s & vbLf
    s = s & "Only put the designated element as such into the table lines. The designated element is not an entire phrase, but is a) only one word or b) a few words in case of a  compound noun, or c) qualifier + word or compound noun. For method steps, the designated element is only the verbal noun as such (e.g. ""obtaining"", ad NOT ""obtaining data from a remote server"")." & vbLf
    s = s & vbLf
    s = s & "## Further observations" & vbLf
    s = s & "Write further observations below the table." & vbLf
    s = s & "Write the further observations in the language which the patent description is written." & vbLf
    s = s & "Be concise when writing further observations." & vbLf
    s = s & "All further observations must be specific to issues you found in the given patent description that relate to reference sign usage" & vbLf
    s = s & vbLf
    s = s & "Include specific further observations that resolve actual naming ambiguities, synonym conflicts, context-dependent sign selection, or other difficulties the paralegal who only inserts reference sign should be aware of." & vbLf
    s = s & vbLf
    s = s & "For cases of multiple reference signs having the same or very similar designed elements, write a concise indication in which claim or in which context which of the ambiguous reference signs must be used, as opposed to other reference signs for the same or a similar designated element. Keep it very short (""use 20 in claims 4 and 5"", ""use 30 collectively for all devices"", ""use 31 specifically for the device when ..."", ""Use S210 for the step of obtaining a user input and S220 for the step of obtaining metadata from a remote server"", ""use 4 for the network interface of the server and 5 for the network interface of the client"" etc.)." & vbLf
    s = s & vbLf
    s = s & "Never explain how to use the table as such." & vbLf
    s = s & vbLf
    s = s & "Never give procedural, mapping workflow, or apparatus-vs-method warnings or other strategic attorney advice." & vbLf
    s = s & vbLf
    s = s & "Never give generic advice on how to insert reference signs unter European practice or on what to do with the table (the paralegal knows that)." & vbLf
    s = s & vbLf
    s = s & "Never mindlessly duplicate information that is directly available in the markdown table and can be retrieved therefrom without any difficulty or ambiguity." & vbLf
    s = s & vbLf
    s = s & "Do not write fluff or general wisdom about claim drafting strategy, reference sign mapping strategy, validity, scope, patent practice, or amendment strategies." & vbLf
    s = s & vbLf
    s = s & vbLf
    s = s & "# Output example template" & vbLf
    s = s & "```" & vbLf
    s = s & "1" & vbTab & "vehicle" & vbLf
    s = s & "5" & vbTab & "motor" & vbLf
    s = s & "10" & vbTab & "control device" & vbLf
    s = s & "20" & vbTab & "control device" & vbLf
    s = s & "S20" & vbTab & "controlling" & vbLf
    s = s & vbLf
    s = s & "Use 10 for the control device in claims 1-4 and 9-10." & vbLf
    s = s & "Use 20 for the control device in claims 5-8." & vbLf
    s = s & "S20 is the step of controlling a steering angle of the vehicle." & vbLf
    s = s & "```"

    DEF_PromptPopulate = s
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
' JSON SERIALIZATION AND DESERIALIZATION
'=======================================================================

Public Function VariablesToJsonObjectString() As String
    Dim json As String
    
    json = "{" & Chr$(34) & "version" & Chr$(34) & ": " & Chr$(34) & TOOL_VERSION & Chr$(34) & "," & vbCrLf
    json = json & "  " & Chr$(34) & "settings" & Chr$(34) & ": {" & vbCrLf
    
    ' String settings
    json = json & "    " & Chr$(34) & "ApiUrl" & Chr$(34) & ": " & Chr$(34) & JsonEscape(gApiUrl) & Chr$(34) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "ApiKey" & Chr$(34) & ": " & Chr$(34) & JsonEscape(gApiKey) & Chr$(34) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "ModelName" & Chr$(34) & ": " & Chr$(34) & JsonEscape(gModelName) & Chr$(34) & "," & vbCrLf
    
    ' Numeric settings (JSON accepts standard decimal point, independent of local format)
    json = json & "    " & Chr$(34) & "Temperature" & Chr$(34) & ": " & FormatDotDouble(gTemperature) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "TimeoutSec" & Chr$(34) & ": " & CStr(gTimeoutSec) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "MaxTokens" & Chr$(34) & ": " & CStr(gMaxTokens) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "Thinking" & Chr$(34) & ": " & IIf(gThinking, "true", "false") & "," & vbCrLf
    
    ' Population settings
    json = json & "    " & Chr$(34) & "TempPopulate" & Chr$(34) & ": " & FormatDotDouble(gTempPopulate) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "TimeoutSecPopulate" & Chr$(34) & ": " & CStr(gTimeoutSecPopulate) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "ThinkPopulation" & Chr$(34) & ": " & IIf(gThinkPopulation, "true", "false") & "," & vbCrLf
    
    ' Other settings
    json = json & "    " & Chr$(34) & "Debug" & Chr$(34) & ": " & IIf(gDebug, "true", "false") & "," & vbCrLf
    
    ' Prompt strings (these contain the heavy lifting for special characters)
    json = json & "    " & Chr$(34) & "PromptInsert" & Chr$(34) & ": " & Chr$(34) & JsonEscape(gPromptInsert) & Chr$(34) & "," & vbCrLf
    json = json & "    " & Chr$(34) & "PromptPopulate" & Chr$(34) & ": " & Chr$(34) & JsonEscape(gPromptPopulate) & Chr$(34) & vbCrLf
    
    json = json & "  }" & vbCrLf & "}"
    
    VariablesToJsonObjectString = json
End Function

' ======================================================================
' JSON SERIALIZATION AND DESERIALIZATION (Continued)
' ======================================================================

' Helper to extract the version string from JSON root level for validation
Private Function ExtractVersionFromJson(jsonText As String) As String
    Dim pKeyStart As Long      ' Position of opening quote of "version" key
    Dim pKeyEnd As Long        ' Position of closing quote of "version" key
    Dim pColon As Long         ' Position of colon after the key
    Dim pValueStart As Long    ' Position of opening quote of version value
    Dim pValueEnd As Long      ' Position of closing quote of version value
    
    ' Find "version" key at root level (case-sensitive)
    pKeyStart = InStr(1, jsonText, Chr$(34) & "version" & Chr$(34), vbBinaryCompare)
    If pKeyStart = 0 Then Exit Function
    
    ' Key end position is pKeyStart + length of '"version"' (9 chars)
    pKeyEnd = pKeyStart + 8
    
    ' Find colon after the key, starting from immediately after closing quote
    pColon = InStr(pKeyEnd + 1, jsonText, ":")
    If pColon = 0 Then Exit Function
    
    ' Skip whitespace to find the opening quote of the value
    Do While pColon <= Len(jsonText) And (Mid$(jsonText, pColon, 1) = " " Or Mid$(jsonText, pColon, 1) = vbTab)
        pColon = pColon + 1
    Loop
    
    If pColon >= Len(jsonText) Then Exit Function
    
    pValueStart = InStr(pColon, jsonText, Chr$(34), vbBinaryCompare)
    If pValueStart = 0 Then Exit Function
    
    ' Find the matching closing quote (handles escaped quotes inside the string)
    pValueEnd = FindJsonStringEnd(jsonText, pValueStart + 1)
    If pValueEnd = 0 Then Exit Function
    
    ExtractVersionFromJson = Mid$(jsonText, pValueStart + 1, pValueEnd - pValueStart - 1)
End Function

' Extracts the major.minor prefix from a version string (e.g., "0.2.0-beta" -> "0.2")
Private Function GetVersionMajorMinor(versionStr As String) As String
    Dim firstDot As Long
    Dim secondDot As Long
    
    ' Find first dot (separates major from minor)
    firstDot = InStr(1, versionStr, ".")
    If firstDot = 0 Then
        ' No dots found - return whole string as fallback
        GetVersionMajorMinor = versionStr
        Exit Function
    End If
    
    ' Find second dot (separates minor from patch)
    secondDot = InStr(firstDot + 1, versionStr, ".")
    If secondDot = 0 Then
        ' Only one dot found - return everything up to end of string (major.minor or major only)
        GetVersionMajorMinor = Left$(versionStr, firstDot)  ' Include the dot but not what follows
        Exit Function
    End If
    
    ' Return everything from start to just before the second dot (i.e., major.minor)
    GetVersionMajorMinor = Left$(versionStr, secondDot - 1)
End Function

' Returns True if version1 is strictly older than version2 (based on major.minor comparison)
Private Function IsVersionOlderThan(version1 As String, version2 As String) As Boolean
    Dim major1 As Long, minor1 As Long
    Dim major2 As Long, minor2 As Long
    Dim firstDot As Long, secondDot As Long
    Dim rest As String
    
    ' Parse version1: major.minor
    major1 = 0: minor1 = 0
    firstDot = InStr(1, version1, ".")
    If firstDot > 0 Then
        major1 = CLng(Left$(version1, firstDot - 1))
        secondDot = InStr(firstDot + 1, version1, ".")
        If secondDot > 0 Then
            minor1 = CLng(Mid$(version1, firstDot + 1, secondDot - firstDot - 1))
        Else
            ' No second dot - try to parse everything after first dot as minor
            rest = Mid$(version1, firstDot + 1)
            If InStr(rest, "-") > 0 Then rest = Left$(rest, InStr(rest, "-") - 1)  ' Remove pre-release suffix
            If IsNumeric(rest) Then minor1 = CLng(rest)
        End If
    Else
        major1 = CLng(version1)
    End If
    
    ' Parse version2: major.minor
    major2 = 0: minor2 = 0
    firstDot = InStr(1, version2, ".")
    If firstDot > 0 Then
        major2 = CLng(Left$(version2, firstDot - 1))
        secondDot = InStr(firstDot + 1, version2, ".")
        If secondDot > 0 Then
            minor2 = CLng(Mid$(version2, firstDot + 1, secondDot - firstDot - 1))
        Else
            rest = Mid$(version2, firstDot + 1)
            If InStr(rest, "-") > 0 Then rest = Left$(rest, InStr(rest, "-") - 1)
            If IsNumeric(rest) Then minor2 = CLng(rest)
        End If
    Else
        major2 = CLng(version2)
    End If
    
    ' Compare: older if major is smaller, or same major but minor is smaller
    If major1 < major2 Then
        IsVersionOlderThan = True
    ElseIf major1 = major2 And minor1 < minor2 Then
        IsVersionOlderThan = True
    Else
        IsVersionOlderThan = False
    End If
End Function

Public Function LoadVariablesFromJsonObjectString(jsonText As String, _
                                                  Optional ByVal bypassVersionCheck As Boolean = False) As Boolean
    Dim settingsStart As Long
    Dim settingsEnd As Long
    Dim scanArea As String
    Dim loadedVersion As String
    
    ' ERROR HANDLING: Save original values for rollback on failure
    Dim orig_ApiUrl As String
    Dim orig_ApiKey As String
    Dim orig_ModelName As String
    Dim orig_Temperature As Double
    Dim orig_TimeoutSec As Long
    Dim orig_MaxTokens As Long
    Dim orig_Thinking As Boolean
    Dim orig_TempPopulate As Double
    Dim orig_TimeoutSecPopulate As Long
    Dim orig_ThinkPopulation As Boolean
    Dim orig_Debug As Boolean
    Dim orig_PromptInsert As String
    Dim orig_PromptPopulate As String
    
    ' Capture current values before any modification
    orig_ApiUrl = gApiUrl
    orig_ApiKey = gApiKey
    orig_ModelName = gModelName
    orig_Temperature = gTemperature
    orig_TimeoutSec = gTimeoutSec
    orig_MaxTokens = gMaxTokens
    orig_Thinking = gThinking
    orig_TempPopulate = gTempPopulate
    orig_TimeoutSecPopulate = gTimeoutSecPopulate
    orig_ThinkPopulation = gThinkPopulation
    orig_Debug = gDebug
    orig_PromptInsert = gPromptInsert
    orig_PromptPopulate = gPromptPopulate
    
    On Error GoTo LoadFailed
    
    ' Validate version compatibility first (skip if bypass requested)
    If Not bypassVersionCheck Then
        loadedVersion = ExtractVersionFromJson(jsonText)
        If Len(loadedVersion) > 0 Then
            ' Compare only major.minor versions (ignore patch and build numbers)
            If GetVersionMajorMinor(loadedVersion) <> GetVersionMajorMinor(TOOL_VERSION) Then
                ' Version mismatch - return False but do not raise error
                LoadVariablesFromJsonObjectString = False
                GoTo RollbackAndExit
            End If
        End If
    End If
    
    ' Find the "settings" object
    settingsStart = InStr(1, jsonText, Chr$(34) & "settings" & Chr$(34), vbBinaryCompare)
    If settingsStart = 0 Then Err.Raise 5, , "Missing required 'settings' key in JSON. " & _
            "Settings file may be corrupted or from an incompatible version."
    
    settingsStart = InStr(settingsStart + 10, jsonText, "{")
    If settingsStart = 0 Then Err.Raise 5, , "Malformed JSON object structure." & _
            " Settings file may be corrupted."
    
    settingsEnd = FindMatchingBracket(jsonText, settingsStart)
    If settingsEnd = 0 Then Err.Raise 5, , "Could not find closing bracket for 'settings'." & _
            " Settings file may be corrupted."
    
    scanArea = Mid$(jsonText, settingsStart + 1, settingsEnd - settingsStart - 1)
    
    ' Extract all values (validation errors will trigger rollback below)
    gApiUrl = ExtractStringKey(scanArea, "ApiUrl")
    gApiKey = ExtractStringKey(scanArea, "ApiKey")
    gModelName = ExtractStringKey(scanArea, "ModelName")
    gTemperature = ExtractNumberKeyDouble(scanArea, "Temperature")
    gTimeoutSec = CLng(ExtractNumberKeyLong(scanArea, "TimeoutSec"))
    gMaxTokens = CLng(ExtractNumberKeyLong(scanArea, "MaxTokens"))
    gThinking = ExtractBooleanKey(scanArea, "Thinking")
    gTempPopulate = ExtractNumberKeyDouble(scanArea, "TempPopulate")
    gTimeoutSecPopulate = CLng(ExtractNumberKeyLong(scanArea, "TimeoutSecPopulate"))
    gThinkPopulation = ExtractBooleanKey(scanArea, "ThinkPopulation")
    gDebug = ExtractBooleanKey(scanArea, "Debug")
    gPromptInsert = ExtractStringKey(scanArea, "PromptInsert")
    gPromptPopulate = ExtractStringKey(scanArea, "PromptPopulate")
    
    ' Success - exit normally without rollback
    LoadVariablesFromJsonObjectString = True
    Exit Function
    
LoadFailed:
    ' ROLLBACK: Restore all original values on any error
    gApiUrl = orig_ApiUrl
    gApiKey = orig_ApiKey
    gModelName = orig_ModelName
    gTemperature = orig_Temperature
    gTimeoutSec = orig_TimeoutSec
    gMaxTokens = orig_MaxTokens
    gThinking = orig_Thinking
    gTempPopulate = orig_TempPopulate
    gTimeoutSecPopulate = orig_TimeoutSecPopulate
    gThinkPopulation = orig_ThinkPopulation
    gDebug = orig_Debug
    gPromptInsert = orig_PromptInsert
    gPromptPopulate = orig_PromptPopulate
    
RollbackAndExit:
    LoadVariablesFromJsonObjectString = False
End Function

' Helper functions for the deserializer
Private Function ExtractStringKey(scanArea As String, keyName As String) As String
    Dim pKey As Long
    Dim pColon As Long
    Dim pOpenQuote As Long
    Dim pCloseQuote As Long
    
    pKey = InStr(1, scanArea, Chr$(34) & keyName & Chr$(34))
    If pKey = 0 Then Err.Raise 5, , "Missing required key: " & keyName & _
            ". Settings file may be corrupted or from an incompatible version."
    
    pColon = InStr(pKey + Len(keyName) + 2, scanArea, ":")
    If pColon = 0 Then Err.Raise 5, , "Missing colon for key: " & keyName
    
    ' Skip whitespace and quotes to find the string value start
    pOpenQuote = InStr(pColon, scanArea, Chr$(34))
    If pOpenQuote = 0 Then Err.Raise 5, , "Malformed string value for key: " & keyName
    
    pCloseQuote = FindJsonStringEnd(scanArea, pOpenQuote + 1)
    If pCloseQuote = 0 Then Err.Raise 5, , "Unterminated string for key: " & keyName
    
    ExtractStringKey = JsonUnescape(Mid$(scanArea, pOpenQuote + 1, pCloseQuote - pOpenQuote - 1))
End Function

Private Function ExtractNumberKeyDouble(scanArea As String, keyName As String) As Double
    Dim pKey As Long
    Dim pColon As Long
    Dim i As Long
    Dim ch As String
    Dim numStr As String
    
    pKey = InStr(1, scanArea, Chr$(34) & keyName & Chr$(34))
    If pKey = 0 Then Err.Raise 5, , "Missing required key: " & keyName & _
            ". Settings file may be corrupted or from an incompatible version."
    
    pColon = InStr(pKey + Len(keyName) + 2, scanArea, ":")
    If pColon = 0 Then Err.Raise 5, , "Missing colon for key: " & keyName
    
    ' Find the start of the number (skip whitespace after colon)
    i = pColon + 1
    Do While i <= Len(scanArea) And (Mid$(scanArea, i, 1) = " " Or Mid$(scanArea, i, 1) = vbTab Or Mid$(scanArea, i, 1) = vbCr Or Mid$(scanArea, i, 1) = vbLf)
        i = i + 1
    Loop
    
    ' Extract number characters
    numStr = ""
    Do While i <= Len(scanArea)
        ch = Mid$(scanArea, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
            numStr = numStr & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    ' Validate extracted number string
    If Len(numStr) = 0 Then Err.Raise 5, , "Empty number value for key: " & keyName & _
            ". Settings file may be corrupted."
    If Not IsNumeric(numStr) Then Err.Raise 5, , "Invalid number format '" & numStr & _
            "' for key: " & keyName & ". Settings file may be corrupted."
    
    ' Use ParseDotDouble for locale-independent decimal point handling
    ExtractNumberKeyDouble = ParseDotDouble(numStr)
End Function

Private Function ExtractNumberKeyLong(scanArea As String, keyName As String) As String
    Dim pKey As Long
    Dim pColon As Long
    Dim i As Long
    Dim ch As String
    Dim numStr As String
    
    pKey = InStr(1, scanArea, Chr$(34) & keyName & Chr$(34))
    If pKey = 0 Then Err.Raise 5, , "Missing required key: " & keyName & _
            ". Settings file may be corrupted or from an incompatible version."
    
    pColon = InStr(pKey + Len(keyName) + 2, scanArea, ":")
    If pColon = 0 Then Err.Raise 5, , "Missing colon for key: " & keyName
    
    i = pColon + 1
    Do While i <= Len(scanArea) And (Mid$(scanArea, i, 1) = " " Or Mid$(scanArea, i, 1) = vbTab Or Mid$(scanArea, i, 1) = vbCr Or Mid$(scanArea, i, 1) = vbLf)
        i = i + 1
    Loop
    
    numStr = ""
    Do While i <= Len(scanArea)
        ch = Mid$(scanArea, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Then
            numStr = numStr & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    ' Validate extracted number string
    If Len(numStr) = 0 Then Err.Raise 5, , "Empty number value for key: " & keyName & _
            ". Settings file may be corrupted."
    If Not IsNumeric(numStr) Then Err.Raise 5, , "Invalid integer format '" & numStr & _
            "' for key: " & keyName & ". Settings file may be corrupted."
    
    ExtractNumberKeyLong = numStr
End Function

Private Function ExtractBooleanKey(scanArea As String, keyName As String) As Boolean
    Dim pKey As Long
    Dim pColon As Long
    Dim i As Long
    Dim ch As String
    Dim valStr As String
    
    pKey = InStr(1, scanArea, Chr$(34) & keyName & Chr$(34))
    If pKey = 0 Then Err.Raise 5, , "Missing required key: " & keyName & _
            ". Settings file may be corrupted or from an incompatible version."
    
    pColon = InStr(pKey + Len(keyName) + 2, scanArea, ":")
    If pColon = 0 Then Err.Raise 5, , "Missing colon for key: " & keyName
    
    i = pColon + 1
    Do While i <= Len(scanArea) And (Mid$(scanArea, i, 1) = " " Or Mid$(scanArea, i, 1) = vbTab Or Mid$(scanArea, i, 1) = vbCr Or Mid$(scanArea, i, 1) = vbLf)
        i = i + 1
    Loop
    
    valStr = ""
    Do While i <= Len(scanArea)
        ch = Mid$(scanArea, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Then
            valStr = valStr & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    ExtractBooleanKey = (LCase$(valStr) = "true")
End Function

'=======================================================================
' PERSISTENCE
'=======================================================================

Public Sub LoadPatentToolsSettings()
    Dim jsonText As String
    Dim loadSuccess As Boolean
    Dim loadedVersion As String
    Dim versionsIncompatible As Boolean
    
    ' Load consolidated settings JSON from registry (single key)
    jsonText = GetSetting(APP_NAME, SECTION_NAME, "SettingsJson", "")
    
    If Len(jsonText) > 0 Then
        ' Extract version first to decide how to handle incompatibility
        loadedVersion = ExtractVersionFromJson(jsonText)
        versionsIncompatible = False
        
        If Len(loadedVersion) > 0 Then
            ' Check for major.minor incompatibility
            If GetVersionMajorMinor(loadedVersion) <> GetVersionMajorMinor(TOOL_VERSION) Then
                versionsIncompatible = True
            End If
        End If
        
        If Not versionsIncompatible Then
            ' Versions compatible (same major.minor, patch may differ) - try normal load
            loadSuccess = LoadVariablesFromJsonObjectString(jsonText, False)
            
            If Not loadSuccess Then
                ' Parse error or other issue - revert to defaults silently
                Call LoadDefaultsToVariables()
                Call SavePatentToolsSettings
            End If
        Else
            ' Versions incompatible - decide based on age
            If IsVersionOlderThan(loadedVersion, TOOL_VERSION) Then
                ' Settings are OLDER than current tool - offer user choice to try loading anyway
                Dim userChoice As VbMsgBoxResult
                userChoice = MsgBox( _
                    "Your saved settings are from an older version of PatentTools." & vbCrLf & vbCrLf & _
                    "Try to use saved settings anyway?", _
                    vbYesNo + vbQuestion, "Settings Version Older")
                
                If userChoice = vbYes Then
                    ' User chose to try loading old settings - bypass version check
                    loadSuccess = LoadVariablesFromJsonObjectString(jsonText, True)
                    If Not loadSuccess Then
                        ' Loading failed - fall back to defaults
                        Call LoadDefaultsToVariables()
                        Call SavePatentToolsSettings
                    End If
                Else
                    ' User chose no - load defaults directly
                    Call LoadDefaultsToVariables()
                    Call SavePatentToolsSettings
                End If
            Else
                ' Settings are NEWER than current tool - silently use defaults, no dialog
                Call LoadDefaultsToVariables()
                Call SavePatentToolsSettings
            End If
        End If
    Else
        ' No settings found: load factory defaults
        Call LoadDefaultsToVariables()
        
        ' Persist defaults so they survive for future loads
        Call SavePatentToolsSettings
    End If
End Sub

' Resets all globally persisted settings in the Windows registry to their factory defaults.
Public Sub ResetSettingsToDefaults()
    ' Load defaults into variables
    Call LoadDefaultsToVariables()
    
    ' Persist to registry
    Call SavePatentToolsSettings
End Sub

' Loads factory-default values into all module-level variables (without persisting to registry).
Private Sub LoadDefaultsToVariables()
    gApiUrl           = DEF_ApiUrl
    gApiKey           = DEF_ApiKey
    gModelName        = DEF_ModelName
    gTemperature      = DEF_Temperature
    gTimeoutSec       = DEF_TimeoutSec
    gMaxTokens        = DEF_MaxTokens
    gThinking         = DEF_Thinking
    gTempPopulate     = DEF_TempPopulate
    gTimeoutSecPopulate = DEF_TimeoutSecPopulate
    gThinkPopulation  = DEF_ThinkPopulation
    gDebug            = DEF_Debug
    gPromptInsert     = DEF_PromptInsert()
    gPromptPopulate   = DEF_PromptPopulate()
End Sub

Public Sub SavePatentToolsSettings()
    Dim jsonText As String
    
    ' Serialize all module-level variables to JSON
    jsonText = VariablesToJsonObjectString()
    
    ' Save consolidated JSON to registry (single key)
    SaveSetting APP_NAME, SECTION_NAME, "SettingsJson", jsonText
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



