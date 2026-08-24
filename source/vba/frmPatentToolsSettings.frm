VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPatentToolsSettings    Caption         =   "PatentTools Settings"   ClientHeight    =   6090   ClientLeft      =   110   ClientTop       =   460   ClientWidth     =   9580.001   OleObjectBlob   =   "frmPatentToolsSettings.frx":0000   StartUpPosition =   1  'FenstermitteEnd
Attribute VB_Name = "frmPatentToolsSettings"Attribute VB_GlobalNameSpace = FalseAttribute VB_Creatable = FalseAttribute VB_PredeclaredId = TrueAttribute VB_Exposed = False
Option Explicit
Private mCancelled As BooleanPrivate mModelsFetched As Boolean
Public Property Get Cancelled() As Boolean    Cancelled = mCancelledEnd Property
Private Sub btnCancel_Click()    mCancelled = True    Me.HideEnd Sub

Private Sub UserForm_Activate()    ' Reset status label whenever the form is shown, so stale connection messages don't persist.    lblStatus.Caption = ""    lblStatus.ForeColor = RGB(255, 165, 0)   ' default yellow for "ready"End SubPrivate Sub AcceptAndSaveChangedSettings()    gApiUrl = NormalizeApiBaseUrl(Trim$(txtApiUrl.Text))    gApiKey = txtApiKey.Text    gModelName = Trim$(cbbModel.Text)    gTemperature = ParseDotDouble(Trim$(txtTemperature.Text))    gTimeoutSec = CLng(Trim$(txtTimeoutSec.Text))    gMaxTokens = CLng(Trim$(txtMaxTokens.Text))    gThinking = chkThinking.Value    gDebug = chkDebug.Value     ' System prompt sent by modRefSigns; stored with vbLf line breaks.    gPromptInsert = FromDisplayText(txtPromptInsert.Text)    SavePatentToolsSettingsEnd Sub
Private Sub btnOK_Click()
    If Not ValidateInputs() Then Exit Sub       AcceptAndSaveChangedSettings
    mCancelled = False
    Me.Hide
End SubPrivate Sub cmdExportSettings_Click()    'Before exporting, validate and save    If Not ValidateInputs() Then Exit Sub   	AcceptAndSaveChangedSettings		'---	'insert code for exporting to file	MsgBox "Not yet implemented.", vbExclamation, "Patent Tools"	'---	End SubPrivate Sub cmdImportSettings_Click()    '---	'insert code for importing from file. 	'we do not persist the imported values, that will only happen after OK, because they could be invalid	MsgBox "Not yet implemented.", vbExclamation, "Patent Tools"	'---		RefreshUiControlsEnd Sub
Private Sub txtApiUrl_Change()
    ' Whenever the URL changes, invalidate any previous successful fetch state.    mModelsFetched = False    btnOK.Enabled = False
    Dim urlText As String
    urlText = Trim$(txtApiUrl.Text)
    If Len(urlText) = 0 Then        lblStatus.ForeColor = RGB(192, 0, 0)        lblStatus.Caption = "URL must not be empty."    Else        lblStatus.ForeColor = RGB(255, 165, 0)        lblStatus.Caption = "Please fetch model list to verify connection."    End If
    ' Clear the model dropdown so stale entries are not presented.    Do While cbbModel.ListCount > 0        cbbModel.RemoveItem (0)    Loop
End Sub
Private Sub txtApiKey_Change()
    ' Whenever the API key changes, invalidate any previous successful fetch state.
    mModelsFetched = False    btnOK.Enabled = False
    lblStatus.ForeColor = RGB(255, 165, 0)   ' yellow/orange    lblStatus.Caption = "Please fetch model list to verify connection."
    ' Clear the model dropdown so stale entries from the previous key are not presented.    Do While cbbModel.ListCount > 0        cbbModel.RemoveItem (0)    Loop
End Sub
Private Sub txtTimeoutSec_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii        Case vbKeyBack        Case Asc("0") To Asc("9")        Case Else            KeyAscii = 0    End Select
End Sub
Private Sub txtMaxTokens_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii        Case vbKeyBack        Case Asc("0") To Asc("9")        Case Else            KeyAscii = 0    End Select
End Sub
Private Sub txtTemperature_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii        Case vbKeyBack        Case Asc("0") To Asc("9")        Case Asc(".")            If InStr(1, Me.txtTemperature.Text, ".") > 0 Then                KeyAscii = 0            End If        Case Else            KeyAscii = 0    End Select
End Sub

Private Function ValidateInputs() As Boolean    Dim s As String
    s = Trim$(txtApiUrl.Text)
    If s = "" Then        MsgBox "OpenAI compatible URL is required.", vbExclamation, "Patent Tools"        txtApiUrl.SetFocus        Exit Function    End If 
    s = Trim$(cbbModel.Text)
    If s = "" Then        MsgBox "Model name is required.", vbExclamation, "Patent Tools"        cbbModel.SetFocus        Exit Function    End If  
    s = Trim$(txtTemperature.Text)    If s = "" Then        MsgBox "Temperature for inserting reference signs is required.", vbExclamation, "Patent Tools"
        txtTemperature.SetFocus
        Exit Function
    End If
    If Not IsDotFloat(s) Then
        MsgBox "Temperature must be a floating-point number using a dot, for example 0.2", vbExclamation, "Patent Tools"
        txtTemperature.SetFocus
        Exit Function
    End If
    s = Trim$(txtTimeoutSec.Text)
    If s = "" Then
        MsgBox "Timeout is required.", vbExclamation, "Patent Tools"
        txtTimeoutSec.SetFocus
        Exit Function
    End If
    If Not IsAllDigits(s) Then
        MsgBox "Timeout must be an integer number of seconds.", vbExclamation, "Patent Tools"
        txtTimeoutSec.SetFocus
        Exit Function
    End If
    s = Trim$(txtMaxTokens.Text)    If s = "" Then        MsgBox "Max. Tokens is required.", vbExclamation, "Patent Tools"        txtMaxTokens.SetFocus        Exit Function    End If
    If Not IsAllDigits(s) Then
        MsgBox "Max. Tokens must be an integer number.", vbExclamation, "Patent Tools"
        txtMaxTokens.SetFocus
        Exit Function
    End If 
    s = Trim$(txtPromptInsert.Text)
    If s = "" Then
        MsgBox "The system prompt must not be empty.", vbExclamation, "Patent Tools"
        ShowPromptInsert
        Exit Function
    End If
    If InStr(1, s, PROMPT_COUNT_TOKEN, vbBinaryCompare) = 0 Then
        If MsgBox("The system prompt no longer contains the placeholder " & PROMPT_COUNT_TOKEN & _
            ", which is replaced with the number of paragraphs sent to the model." & vbCrLf & vbCrLf & _
            "The required paragraph count will then be appended to the prompt automatically." & vbCrLf & vbCrLf & _			"Not informing the model about how much paragraphs it is to output will decrease stability." & vbCrLf & vbCrLf & _
            "Save anyway?", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Patent Tools") <> vbYes Then            ShowPromptInsert            Exit Function        End If
    End If
    ValidateInputs = True
End Function
Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    s = Trim$(s)    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)        ch = Mid$(s, i, 1)        If ch < "0" Or ch > "9" Then Exit Function    Next i 
    IsAllDigits = True
End Function
Private Function IsDotFloat(ByVal s As String) As Boolean    Dim i As Long    Dim ch As String    Dim dotCount As Long    Dim digitCount As Long        s = Trim$(s)    If Len(s) = 0 Then Exit Function        For i = 1 To Len(s)        ch = Mid$(s, i, 1)        Select Case ch            Case "0" To "9"                digitCount = digitCount + 1            Case "."                dotCount = dotCount + 1                If dotCount > 1 Then Exit Function            Case Else                Exit Function        End Select    Next i        If digitCount = 0 Then Exit Function        IsDotFloat = TrueEnd Function
Private Sub UserForm_Initialize()    mCancelled = True	mpTabs.Value = 0  ' Erste Registerkarte aktivieren        LoadPatentToolsSettings  
    ' Load current global variables into the UI.
    RefreshUIControls
    ' Initial state: OK enabled, status label empty.
    lblStatus.Caption = ""
    mModelsFetched = True  ' persisted model counts as confirmed
    btnOK.Enabled = True
End Sub
'-------------------------------------------------------------------
' Refreshes all settings controls based on the currently loaded globals.
' This keeps the code DRY and consistent across Initialize and Reset.
'-------------------------------------------------------------------
Private Sub RefreshUIControls()
    txtApiUrl.Text = gApiUrl
    txtApiKey.Text = gApiKey
    ' Start with persisted model selected so OK is usable without network access.    Do While cbbModel.ListCount > 0        cbbModel.RemoveItem (0)
    Loop
    If Trim$(gModelName) <> "" Then
        cbbModel.AddItem (gModelName)
        cbbModel.ListIndex = 0
    End If        txtTemperature.Text = FormatDotDouble(gTemperature)    txtTimeoutSec.Text = CStr(gTimeoutSec)    txtMaxTokens.Text = CStr(gMaxTokens)    chkThinking.Value = gThinking    chkDebug.Value = gDebug        ' System prompt editor: multi-line and scrollable, showing the persisted prompt.    With txtPromptInsert        .MultiLine = True        .EnterKeyBehavior = True        .WordWrap = True        .ScrollBars = fmScrollBarsVertical        .Text = ToDisplayText(gPromptInsert)        .SelStart = 0    End WithEnd Sub

Private Sub FetchModelsFromAPI()
    Dim modelNames As Collection
    Dim errText As String
    Dim i As Long

    ' Explicitly instantiate the collection to avoid ByRef pass-by-Nothing issues.
    Set modelNames = New Collection
    errText = ""

    On Error Resume Next
    If FetchModelList(Trim$(txtApiUrl.Text), modelNames, errText) Then
        On Error GoTo 0

        ' Success: replace dropdown with server list and select first item.
        Do While cbbModel.ListCount > 0
            cbbModel.RemoveItem (0)
        Loop
        For i = 1 To modelNames.Count
            cbbModel.AddItem (modelNames(i))
        Next i
        cbbModel.ListIndex = 0

        lblStatus.ForeColor = RGB(0, 128, 0)
        lblStatus.Caption = "Connection successful."
        mModelsFetched = True
        btnOK.Enabled = True
    Else
        On Error GoTo 0

        ' Failure: keep persisted values but warn the user.
        lblStatus.ForeColor = RGB(192, 0, 0)
        lblStatus.Caption = "Connection failed: " & errText

        ' OK stays enabled so the user can confirm the persisted config anyway.
        mModelsFetched = True
        btnOK.Enabled = True
    End If
End Sub

Private Sub btnRefreshModels_Click()
    FetchModelsFromAPI
End Sub


' Brings txtPromptInsert into view (it may live on another MultiPage tab) and focuses it.
Private Sub ShowPromptInsert()
    On Error Resume Next
    mpTabs.Value = txtPromptInsert.Parent.Index
    txtPromptInsert.SetFocus
    On Error GoTo 0
End Sub
Private Sub cmdResetToDefault_Click()
    If MsgBox("Are you sure you want to reset all settings to defaults?", vbYesNo + vbQuestion, "Patent Tools") = vbYes Then
        ResetSettingsToDefaults
        RefreshUIControls
    End If
End Sub
