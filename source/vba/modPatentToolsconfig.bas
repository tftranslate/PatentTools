Attribute VB_Name = "modPatentToolsconfig"
Option Explicit

Public gApiUrl As String
Public gApiKey As String
Public gModelName As String
Public gTemperature As Double
Public gTimeoutSec As Long
Public gMaxTokens As Long
Public gThinking As Boolean
Public gDebug As Boolean

Private Const APP_NAME As String = "PatentTools"
Private Const SECTION_NAME As String = "Settings"

Public Sub LoadPatentToolsSettings()
    gApiUrl = GetSetting(APP_NAME, SECTION_NAME, "ApiUrl", "https://api.openai.com/v1")
    gApiKey = GetSetting(APP_NAME, SECTION_NAME, "ApiKey", "")
    gModelName = GetSetting(APP_NAME, SECTION_NAME, "ModelName", "gpt-4o")
    gTemperature = CDbl(GetSetting(APP_NAME, SECTION_NAME, "Temperature", "0.2"))
    gTimeoutSec = CLng(GetSetting(APP_NAME, SECTION_NAME, "TimeoutSec", "120"))
    gMaxTokens = CLng(GetSetting(APP_NAME, SECTION_NAME, "MaxTokens", "32768"))
    gThinking = CBool(Val(GetSetting(APP_NAME, SECTION_NAME, "Thinking", "0")))
    gDebug = CBool(Val(GetSetting(APP_NAME, SECTION_NAME, "Debug", "0")))
End Sub

Public Sub SavePatentToolsSettings()
    SaveSetting APP_NAME, SECTION_NAME, "ApiUrl", gApiUrl
    SaveSetting APP_NAME, SECTION_NAME, "ApiKey", gApiKey
    SaveSetting APP_NAME, SECTION_NAME, "ModelName", gModelName
    SaveSetting APP_NAME, SECTION_NAME, "Temperature", Replace(CStr(gTemperature), ",", ".")
    SaveSetting APP_NAME, SECTION_NAME, "TimeoutSec", CStr(gTimeoutSec)
    SaveSetting APP_NAME, SECTION_NAME, "MaxTokens", CStr(gMaxTokens)
    SaveSetting APP_NAME, SECTION_NAME, "Thinking", IIf(gThinking, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "Debug", IIf(gDebug, "1", "0")
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
