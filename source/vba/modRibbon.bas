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
