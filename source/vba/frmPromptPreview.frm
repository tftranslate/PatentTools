VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPromptPreview 
   Caption         =   "Model Call Prompt Preview"
   ClientHeight    =   10110
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   6030
   OleObjectBlob   =   "frmPromptPreview.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmPromptPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Model Call Preview"
    
    With Me.txtPreview
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .Locked = True
        .SelStart = 0
    End With
    
    Cancelled = True
End Sub

Private Sub cmdOK_Click()
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub
