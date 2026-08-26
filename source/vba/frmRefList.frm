VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRefList 
   Caption         =   "Reference Sign List"
   ClientHeight    =   8830.001
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   8660.001
   OleObjectBlob   =   "frmRefList.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRefList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Public Cancelled As Boolean

Private Sub cmdClear_Click()
    ' Clear the reference list text box and remove any persisted value,
    ' so the hardcoded defaults will be used next time this dialog opens.
    txtRefList.Text = ""

    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("PatentToolsRefList").Delete
    On Error GoTo 0
End Sub

Private Sub cmdPopulate_Click()
    Dim refTable As String

    ' Call the model-assisted population routine.
    refTable = Populate_Reference_Sign_Table()

    ' Persist the result document-wide (per-document custom property).
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("PatentToolsRefList").Value = refTable
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="PatentToolsRefList", _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=refTable
    End If
    On Error GoTo 0

    ' Update the text box so the user sees the freshly populated content.
    Me.txtRefList.Text = refTable
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Reference Sign List"

    Dim refList As String
    Dim defaultRefList As String
    
    ' Load persisted reference list from document custom properties, if available.
    On Error Resume Next
    refList = ActiveDocument.CustomDocumentProperties("PatentToolsRefList").Value
    On Error GoTo 0

    If Len(refList) > 0 Then
        ' Persisted value exists: use it.
        defaultRefList = refList
    Else
        ' No persisted value: use hardcoded defaults for first-time users.
        defaultRefList = "vehicle" & vbTab & "10" & vbCrLf & _
                          "wheel" & vbTab & "2" & vbCrLf & _
                          "engine" & vbTab & "3" & vbCrLf & _
                          "surface" & vbTab & "100" & vbCrLf
    End If

    With Me.txtRefList
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .Text = defaultRefList
        .SelStart = 0
    End With

    Cancelled = True
End Sub

Private Sub cmdOK_Click()
    ' Save the reference list to document custom properties for per-document persistence.
    Dim refText As String
    refText = Me.txtRefList.Text

    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("PatentToolsRefList").Value = refText
    If Err.Number <> 0 Then
        ' Property doesn't exist yet: create it.
        ActiveDocument.CustomDocumentProperties.Add Name:="PatentToolsRefList", _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=refText
    End If
    On Error GoTo 0

    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub
