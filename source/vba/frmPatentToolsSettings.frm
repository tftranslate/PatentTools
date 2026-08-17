VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPatentToolsSettings 
   Caption         =   "PatentTools Settings"
   ClientHeight    =   5120
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   9430.001
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

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property


Private Sub btnCancel_Click()
    mCancelled = True
    Me.Hide
End Sub


Private Sub btnOK_Click()
    If Not ValidateInputs() Then Exit Sub
    
    gApiUrl = NormalizeApiBaseUrl(Trim$(txtApiUrl.Text))
    gApiKey = txtApiKey.Text
    gModelName = Trim$(txtModelName.Text)
    gTemperature = CDbl(Replace(Trim$(txtTemperature.Text), ",", "."))
    gTimeoutSec = CLng(Trim$(txtTimeoutSec.Text))
    gMaxTokens = CLng(Trim$(txtMaxTokens.Text))
    gThinking = chkThinking.Value
    gDebug = chkDebug.Value
    
    SavePatentToolsSettings
    
    mCancelled = False
    Me.Hide
End Sub




Private Sub CheckBox1_Click()

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
    
    s = Trim$(txtModelName.Text)
    If s = "" Then
        MsgBox "Model name is required.", vbExclamation, "Patent Tools"
        txtModelName.SetFocus
        Exit Function
    End If
    
    s = Trim$(txtTemperature.Text)
    If s = "" Then
        MsgBox "Temperature is required.", vbExclamation, "Patent Tools"
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
    
    LoadPatentToolsSettings
    
    txtApiUrl.Text = gApiUrl
    txtApiKey.Text = gApiKey
    txtModelName.Text = gModelName
    txtTemperature.Text = Replace(CStr(gTemperature), ",", ".")
    txtTimeoutSec.Text = CStr(gTimeoutSec)
    txtMaxTokens.Text = CStr(gMaxTokens)
    chkThinking.Value = gThinking
    chkDebug.Value = gDebug
End Sub
