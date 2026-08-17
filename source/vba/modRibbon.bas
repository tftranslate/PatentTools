Attribute VB_Name = "modRibbon"
Option Explicit

Public gPatentToolsRibbon As IRibbonUI

Public Sub PatentToolsRibbon_OnLoad(ribbon As IRibbonUI)
    Set gPatentToolsRibbon = ribbon
End Sub

Public Sub PatentTools_InsertReferenceSigns(control As IRibbonControl)
    Insert_Reference_Signs
End Sub

Public Sub PatentTools_OpenSettings(control As IRibbonControl)
    Patent_Tools_Settings
End Sub

Public Sub PatentTools_About(ByVal control As IRibbonControl)
    MsgBox _
        "Patent Tools" & vbCrLf & vbCrLf & _
        "Reference-sign insertion for patent claims." & vbCrLf & vbCrLf & _
        "Copyright (c) 2026 Tobias Ernst" & vbCrLf & _
        "Licensed under the MIT License." & vbCrLf & vbCrLf & _
        "GitHub:" & vbCrLf & _
        "https://github.com/tftranslate/PatentTools", _
        vbInformation, _
        "About Patent Tools"
End Sub


