VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPatentToolsSettings 
   Caption         =   "PatentTools Settings"
   ClientHeight    =   6090
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   9580.001
   OleObjectBlob   =   "frmPatentToolsSettings.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmPatentToolsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private mCancelled As Boolean
Private mModelsFetched As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

' ====================================================================
' HELPERS FOR FILE EXPORT/IMPORT
' ====================================================================

Private Function GetUserDownloadsFolder() As String
    ' Returns empty string - FileDialog will use system default location
    ' This keeps the code platform-independent and simple
    GetUserDownloadsFolder = ""
End Function

Private Function SaveJsonToFile(ByVal jsonContent As String) As Boolean
    Dim saveDialog As FileDialog
    Dim filePath As String
    Dim i As Long
    
    ' Use proper SaveAs dialog with Plain Text filter
    Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
    
    With saveDialog
        .Title = "Export Settings"
        
        ' Find and select the Plain Text (*.txt) filter in the system defaults
        ' This ensures proper text encoding and file type association
        For i = 1 To .Filters.Count
            If InStr(.Filters(i).Description, "Plain Text") > 0 Or _
               InStr(.Filters(i).Extensions, ".txt") > 0 Then
                .FilterIndex = i
                Exit For
            End If
        Next i
        
        ' Let FileDialog use system default location (e.g., Documents or Last used folder)
        .InitialFileName = "PatentTools_Settings.txt"
        
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            
            ' Ensure .txt extension (JSON content is written to plain text file)
            If LCase$(Right$(filePath, 4)) <> ".txt" Then
                filePath = filePath & ".txt"
            End If
            
            ' Write JSON content to plain text file
            Call write_file_binary(filePath, jsonContent)
            
            MsgBox "Settings exported successfully to:" & vbCrLf & filePath, vbInformation, "Patent Tools"
            SaveJsonToFile = True
        End If
    End With
    
    Set saveDialog = Nothing
End Function

Private Function LoadJsonFromFile() As Boolean
    Dim openDialog As FileDialog
    Dim filePath As String
    Dim jsonContent As String
    
    Set openDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With openDialog
        .Title = "Import Settings"
        ' Use .txt filter (settings are stored as plain text with JSON content)
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt", 1
        .Filters.Add "All Files", "*.*", 2
        
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            
            ' Read JSON from file as binary (preserves exact content)
            jsonContent = read_file_binary(filePath)
            
            If Len(jsonContent) > 0 Then
                ' First attempt: strict version check
                Dim loadSuccess As Boolean
                loadSuccess = LoadVariablesFromJsonObjectString(jsonContent, False)
                
                If Not loadSuccess Then
                    ' Version mismatch - ask user if they want to proceed anyway
                    Dim userChoice As VbMsgBoxResult
                    userChoice = MsgBox( _
                        "Settings version mismatch detected." & vbCrLf & vbCrLf & _
                        "Try importing anyway?", _
                        vbYesNo + vbQuestion, "Version Mismatch")
                    
                    If userChoice = vbNo Then
                        LoadJsonFromFile = False
                        Exit Function
                    End If
                    
                    ' User chose to proceed - retry with version check bypassed
                    loadSuccess = LoadVariablesFromJsonObjectString(jsonContent, True)
                    
                    If Not loadSuccess Then
                        MsgBox "Import failed due to JSON parse error.", vbExclamation, "Patent Tools"
                        LoadJsonFromFile = False
                        Exit Function
                    End If
                End If
                
                ' Success path continues...
                MsgBox "Settings imported successfully. Review and click OK to save." & vbCrLf & _
                       "Note: Current session values will be replaced by imported settings.", vbInformation, "Patent Tools"
                LoadJsonFromFile = True
            Else
                MsgBox "Could not read file or file is empty.", vbExclamation, "Patent Tools"
                LoadJsonFromFile = False
            End If
        End If
    End With
    
    Set openDialog = Nothing
End Function

' Write JSON content to file as plain text (using Output mode)
Private Sub write_file_binary(ByVal filePath As String, ByVal content As String)
    Dim fnum As Integer
    fnum = FreeFile
    
    ' Use Output mode for text files - compatible with Print #
    Open filePath For Output As #fnum
    Print #fnum, content
    Close #fnum
End Sub

' Read file content as-is in binary mode
Private Function read_file_binary(ByVal filePath As String) As String
    Dim fnum As Integer
    Dim fileSize As Long
    
    On Error Resume Next
    fnum = FreeFile
    Open filePath For Binary Access Read As #fnum
    fileSize = LOF(fnum)
    
    If fileSize > 0 Then
        read_file_binary = Space$(fileSize)
        Get #fnum, , read_file_binary
    Else
        read_file_binary = ""
    End If
    
    Close #fnum
    On Error GoTo 0
End Function

Private Sub btnCancel_Click()
    mCancelled = True
    Me.Hide
End Sub


Private Sub UserForm_Activate()
    ' Reset status label whenever the form is shown, so stale connection messages don't persist.
    lblStatus.Caption = ""
    lblStatus.ForeColor = RGB(255, 165, 0)   ' default yellow for "ready"
End Sub

Private Sub AcceptAndSaveChangedSettings()

    gApiUrl = NormalizeApiBaseUrl(Trim$(txtApiUrl.Text))
    gApiKey = txtApiKey.Text
    gModelName = Trim$(cbbModel.Text)
    gTemperature = ParseDotDouble(Trim$(txtTemperature.Text))
    gTimeoutSec = CLng(Trim$(txtTimeoutSec.Text))
    gTempPopulate = ParseDotDouble(Trim$(txtTempPopulation.Text))
    gTimeoutSecPopulate = CLng(Trim$(txtTimeoutSecPopulate.Text))
    gThinkPopulation = chkThinkPopulation.Value
    gMaxTokens = CLng(Trim$(txtMaxTokens.Text))
    gThinking = chkThinking.Value
    gDebug = chkDebug.Value
 
    ' System prompts: stored with vbLf line breaks.
    gPromptInsert = FromDisplayText(txtPromptInsert.Text)
    gPromptPopulate = FromDisplayText(txtPromptPopulate.Text)

    SavePatentToolsSettings
End Sub


Private Sub btnOK_Click()

    If Not ValidateInputs() Then Exit Sub
    AcceptAndSaveChangedSettings
    mCancelled = False
    Me.Hide

End Sub

Private Sub cmdExportSettings_Click()
    Dim jsonContent As String

    'Before exporting, validate and save current settings to registry
    If Not ValidateInputs() Then Exit Sub
    AcceptAndSaveChangedSettings
    
    'Generate JSON from current module-level variables
    jsonContent = VariablesToJsonObjectString()
    
    'Export to file (user-selects location in Downloads folder)
    Call SaveJsonToFile(jsonContent)
        
End Sub

Private Sub cmdImportSettings_Click()
    Dim importSuccess As Boolean
    
    'Import from file (loads into variables with rollback on error)
    importSuccess = LoadJsonFromFile()
    
    If importSuccess Then
        'Update dialog display with imported values (not yet persisted to registry)
        RefreshUIControls
        
        'Re-validate imported settings against current server
        'This may grey out OK if connection doesn't work or model not available
        FetchModelsFromAPI
    End If
End Sub


Private Sub txtApiUrl_Change()

    ' Whenever the URL changes, invalidate any previous successful fetch state.
    mModelsFetched = False
    btnOK.Enabled = False

    Dim urlText As String

    urlText = Trim$(txtApiUrl.Text)

    If Len(urlText) = 0 Then
        lblStatus.ForeColor = RGB(192, 0, 0)
        lblStatus.Caption = "URL must not be empty."
    Else
        lblStatus.ForeColor = RGB(255, 165, 0)
        lblStatus.Caption = "Please fetch model list to verify connection."
    End If

    ' Clear the model dropdown so stale entries are not presented.
    Do While cbbModel.ListCount > 0
        cbbModel.RemoveItem (0)
    Loop

End Sub

Private Sub txtApiKey_Change()

    ' Whenever the API key changes, invalidate any previous successful fetch state.

    mModelsFetched = False
    btnOK.Enabled = False

    lblStatus.ForeColor = RGB(255, 165, 0)   ' yellow/orange
    lblStatus.Caption = "Please fetch model list to verify connection."

    ' Clear the model dropdown so stale entries from the previous key are not presented.
    Do While cbbModel.ListCount > 0
        cbbModel.RemoveItem (0)
    Loop

End Sub

Private Sub txtTimeoutSec_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKeyBack
        Case Asc("0") To Asc("9")
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtMaxTokens_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKeyBack
        Case Asc("0") To Asc("9")
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtTemperature_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case vbKeyBack
        Case Asc("0") To Asc("9")
        Case Asc(".")
            If InStr(1, Me.txtTemperature.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select

End Sub


Private Function ValidateInputs() As Boolean
    Dim s As String

    s = Trim$(txtApiUrl.Text)
    If s = "" Then
        MsgBox "OpenAI compatible URL is required.", vbExclamation, "Patent Tools"
        txtApiUrl.SetFocus
        Exit Function
    End If
 
    s = Trim$(cbbModel.Text)
    If s = "" Then
        MsgBox "Model name is required.", vbExclamation, "Patent Tools"
        cbbModel.SetFocus
        Exit Function
    End If
  
    s = Trim$(txtTemperature.Text)
    If s = "" Then
        MsgBox "Temperature for inserting reference signs is required.", vbExclamation, "Patent Tools"
        txtTemperature.SetFocus
        Exit Function
    End If
    If Not IsDotFloat(s) Then
        MsgBox "Temperature must be a floating-point number using a dot, for example 0.2", vbExclamation, "Patent Tools"
        txtTemperature.SetFocus
        Exit Function
    End If

    s = Trim$(txtTempPopulation.Text)
    If s = "" Then
        MsgBox "Temperature for population is required.", vbExclamation, "Patent Tools"
        txtTempPopulation.SetFocus
        Exit Function
    End If
    If Not IsDotFloat(s) Then
        MsgBox "Population Temperature must be a floating-point number using a dot, for example 0.0", vbExclamation, "Patent Tools"
        txtTempPopulation.SetFocus
        Exit Function
    End If

    s = Trim$(txtTimeoutSec.Text)
    If s = "" Then
        MsgBox "Timeout for insertion is required.", vbExclamation, "Patent Tools"
        txtTimeoutSec.SetFocus
        Exit Function
    End If
    If Not IsAllDigits(s) Then
        MsgBox "Timeout must be an integer number of seconds.", vbExclamation, "Patent Tools"
        txtTimeoutSec.SetFocus
        Exit Function
    End If

    s = Trim$(txtTimeoutSecPopulate.Text)
    If s = "" Then
        MsgBox "Timeout for population is required.", vbExclamation, "Patent Tools"
        txtTimeoutSecPopulate.SetFocus
        Exit Function
    End If
    If Not IsAllDigits(s) Then
        MsgBox "Population Timeout must be an integer number of seconds.", vbExclamation, "Patent Tools"
        txtTimeoutSecPopulate.SetFocus
        Exit Function
    End If

    s = Trim$(txtMaxTokens.Text)
    If s = "" Then
        MsgBox "Max. Tokens is required.", vbExclamation, "Patent Tools"
        txtMaxTokens.SetFocus
        Exit Function
    End If

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
            "The required paragraph count will then be appended to the prompt automatically." & vbCrLf & vbCrLf & _
                        "Not informing the model about how much paragraphs it is to output will decrease stability." & vbCrLf & vbCrLf & _
            "Save anyway?", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Patent Tools") <> vbYes Then
            ShowPromptInsert
            Exit Function
        End If
    End If

    s = Trim$(txtPromptPopulate.Text)
    If s = "" Then
        MsgBox "The population system prompt must not be empty.", vbExclamation, "Patent Tools"
        ShowPromptPopulate
        Exit Function
    End If

    ValidateInputs = True
End Function


Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsAllDigits = True

End Function

Private Function IsDotFloat(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim dotCount As Long
    Dim digitCount As Long
    
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "0" To "9"
                digitCount = digitCount + 1
            Case "."
                dotCount = dotCount + 1
                If dotCount > 1 Then Exit Function
            Case Else
                Exit Function
        End Select
    Next i
    
    If digitCount = 0 Then Exit Function
    
    IsDotFloat = True
End Function

Private Sub UserForm_Initialize()
    mCancelled = True
        mpTabs.Value = 0  ' Erste Registerkarte aktivieren
        
    LoadPatentToolsSettings
        
    ' Load current global variables into the UI.
    RefreshUIControls
        
    ' Initial state: OK disabled until validated, status message ready for auto-fetch result
    lblStatus.Caption = ""
    mModelsFetched = False  ' Not yet validated on this machine
    btnOK.Enabled = False   ' Disabled until user validates connection or manually fetches models
    
    ' Auto-fetch model list to validate persisted settings against current server
    FetchModelsFromAPI
End Sub



'-------------------------------------------------------------------
' Refreshes all settings controls based on the currently loaded globals.
' This keeps the code DRY and consistent across Initialize and Reset.
'-------------------------------------------------------------------

Private Sub RefreshUIControls()
    txtApiUrl.Text = gApiUrl
    txtApiKey.Text = gApiKey

    ' Start with persisted model selected so OK is usable without network access.
    Do While cbbModel.ListCount > 0
        cbbModel.RemoveItem (0)
    Loop

    If Trim$(gModelName) <> "" Then
        cbbModel.AddItem (gModelName)
        cbbModel.ListIndex = 0
    End If
    
    txtTemperature.Text = FormatDotDouble(gTemperature)
    txtTimeoutSec.Text = CStr(gTimeoutSec)
    txtTempPopulation.Text = FormatDotDouble(gTempPopulate)
    txtTimeoutSecPopulate.Text = CStr(gTimeoutSecPopulate)
    chkThinkPopulation.Value = gThinkPopulation
    txtMaxTokens.Text = CStr(gMaxTokens)
    chkThinking.Value = gThinking
    chkDebug.Value = gDebug
    
    ' System prompt editor: multi-line and scrollable, showing the persisted prompt.
    With txtPromptInsert
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Text = ToDisplayText(gPromptInsert)
        .SelStart = 0
    End With

    ' Population system prompt editor: multi-line and scrollable, showing the persisted prompt.
    With txtPromptPopulate
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Text = ToDisplayText(gPromptPopulate)
        .SelStart = 0
    End With
End Sub

Private Sub FetchModelsFromAPI()
    Dim modelNames As Collection
    Dim errText As String
    Dim i As Long
    Dim persistedModel As String
    Dim selectedIndex As Long
    Dim modelNameLower As String
    Dim persistLower As String
    
    ' Explicitly instantiate the collection to avoid ByRef pass-by-Nothing issues.
    Set modelNames = New Collection
    errText = ""
    
    ' Save the persisted model name before we clear the dropdown
    persistedModel = Trim$(gModelName)

    On Error Resume Next
    If FetchModelList(Trim$(txtApiUrl.Text), modelNames, errText) Then
        On Error GoTo 0

        ' Success: replace dropdown with server list
        Do While cbbModel.ListCount > 0
            cbbModel.RemoveItem (0)
        Loop
        For i = 1 To modelNames.count
            cbbModel.AddItem (modelNames(i))
        Next i
        
        ' Try to select the persisted model if it exists in the list
        selectedIndex = -1
        If Len(persistedModel) > 0 Then
            persistLower = LCase$(persistedModel)
            For i = 0 To cbbModel.ListCount - 1
                modelNameLower = LCase$(cbbModel.List(i))
                If modelNameLower = persistLower Then
                    selectedIndex = i
                    Exit For
                End If
            Next i
        End If
        
        ' Select matched model or fall back to first item
        If selectedIndex >= 0 Then
            cbbModel.ListIndex = selectedIndex
            lblStatus.ForeColor = RGB(0, 128, 0)
            lblStatus.Caption = "Connection successful. Selected model '" & persistedModel & "'." 
        Else
            cbbModel.ListIndex = 0
            lblStatus.ForeColor = RGB(0, 128, 0)
            If Len(persistedModel) > 0 Then
                lblStatus.Caption = "Connection successful. Persisted model '" & persistedModel & _
                                   "' not found on server; first model selected." 
            Else
                lblStatus.Caption = "Connection successful."
            End If
        End If
        
        mModelsFetched = True
        btnOK.Enabled = True  ' OK enabled after successful validation
    Else
        On Error GoTo 0

        ' Failure: clear model dropdown and warn the user.
        Do While cbbModel.ListCount > 0
            cbbModel.RemoveItem (0)
        Loop
        lblStatus.ForeColor = RGB(192, 0, 0)
        lblStatus.Caption = "Connection failed: " & errText

        ' OK disabled until connection works - user must fix URL/key and retry
        mModelsFetched = False
        btnOK.Enabled = False
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

Private Sub ShowPromptPopulate()
    On Error Resume Next
    mpTabs.Value = txtPromptPopulate.Parent.Index
    txtPromptPopulate.SetFocus
    On Error GoTo 0
End Sub

Private Sub cmdResetToDefault_Click()
    If MsgBox("Are you sure you want to reset all settings to defaults?", vbYesNo + vbQuestion, "Patent Tools") = vbYes Then
        ResetSettingsToDefaults
        RefreshUIControls
    End If
End Sub
