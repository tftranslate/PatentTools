Attribute VB_Name = "modBuildSettingsTabs"
Option Explicit

'=======================================================================
' One-shot design-time restructuring of frmPatentToolsSettings.
'
' Inserts a 3-page MultiPage ("mpTabs") into the form and relocates every
' existing setting control onto page 1, preserving each control's name,
' size, relative position, font, colours and value.
'
' Dialog-level controls (btnOK, btnCancel, lblStatus) stay on the form
' itself, below the tab strip, so they remain visible on all three pages.
'
' Control NAMES are unchanged, therefore the form's existing code-behind
' and all its event handlers (txtApiUrl_Change, btnOK_Click, ...) keep
' working with no edits. Controls nested inside a MultiPage page are
' still exposed as members of the form module in VBA.
'
' PREREQUISITES
'   1. File > Options > Trust Center > Trust Center Settings >
'      Macro Settings > tick "Trust access to the VBA project object model".
'   2. Export a backup of frmPatentToolsSettings first. This is destructive.
'   3. The form must be CLOSED in the designer when you run this.
'
' USAGE: run BuildSettingsTabs once, then Debug > Compile Project.
'=======================================================================

Private Const FORM_NAME      As String = "frmPatentToolsSettings"
Private Const PAGE_INSET     As Single = 10      ' padding inside a page
Private Const TABSTRIP_H     As Single = 22      ' height eaten by the tab row
Private Const OUTER_MARGIN   As Single = 8       ' form edge -> MultiPage
Private Const FOOTER_GAP     As Single = 8       ' MultiPage -> button row

' Controls that must NOT move onto page 1.
Private Const KEEP_ON_FORM   As String = "|btnOK|btnCancel|lblStatus|"

Public Sub BuildSettingsTabs()
    Dim vbc As Object            ' VBComponent
    Dim dsg As Object            ' UserForm designer
    Dim mp As Object             ' MultiPage
    Dim pg As Object
    Dim ctl As Object
    Dim src As Object
    Dim dst As Object

    Dim names() As String
    Dim n As Long, i As Long
    Dim minL As Single, minT As Single, maxR As Single, maxB As Single
    Dim bbW As Single, bbH As Single
    Dim footerTop As Single, formW As Single

    On Error GoTo Fail

    Set vbc = Application.VBE.ActiveVBProject.VBComponents(FORM_NAME)
    Set dsg = vbc.Designer

    If HasControl(dsg, "mpTabs") Then
        MsgBox "mpTabs already exists on " & FORM_NAME & _
               ". Restore your backup before running this again.", _
               vbExclamation, "Patent Tools"
        Exit Sub
    End If

    '-------------------------------------------------------------------
    ' 1. Snapshot the controls to relocate, and their bounding box.
    '-------------------------------------------------------------------
    ReDim names(1 To dsg.Controls.Count)
    minL = 1E+09: minT = 1E+09: maxR = -1E+09: maxB = -1E+09

    For Each ctl In dsg.Controls
        If InStr(1, KEEP_ON_FORM, "|" & ctl.Name & "|", vbTextCompare) = 0 Then
            n = n + 1
            names(n) = ctl.Name
            If ctl.Left < minL Then minL = ctl.Left
            If ctl.Top < minT Then minT = ctl.Top
            If ctl.Left + ctl.Width > maxR Then maxR = ctl.Left + ctl.Width
            If ctl.Top + ctl.Height > maxB Then maxB = ctl.Top + ctl.Height
        End If
    Next ctl

    If n = 0 Then
        MsgBox "No relocatable controls found on " & FORM_NAME & ".", _
               vbExclamation, "Patent Tools"
        Exit Sub
    End If

    bbW = maxR - minL
    bbH = maxB - minT

    '-------------------------------------------------------------------
    ' 2. Create the MultiPage sized to hold the existing block.
    '-------------------------------------------------------------------
    Set mp = dsg.Controls.Add("Forms.MultiPage.1", "mpTabs", True)
    mp.Left = OUTER_MARGIN
    mp.Top = OUTER_MARGIN
    mp.Width = bbW + PAGE_INSET * 2
    mp.Height = bbH + PAGE_INSET * 2 + TABSTRIP_H

    Do While mp.Pages.Count < 3
        mp.Pages.Add
    Loop

    mp.Pages(0).Name = "pgModel"
    mp.Pages(0).Caption = "Model && Connection"
    mp.Pages(1).Name = "pgTab2"
    mp.Pages(1).Caption = "Tab 2"
    mp.Pages(2).Name = "pgTab3"
    mp.Pages(2).Caption = "Tab 3"
    mp.Value = 0

    Set pg = mp.Pages(0)

    '-------------------------------------------------------------------
    ' 3. Recreate each control on page 1, then drop the original.
    '    Rename the original first so the new one can take its name.
    '-------------------------------------------------------------------
    For i = 1 To n
        Set src = dsg.Controls(names(i))
        src.Name = "zzOld_" & names(i)

        Set dst = pg.Controls.Add(ProgIdOf(src), names(i), True)

        dst.Left = src.Left - minL + PAGE_INSET
        dst.Top = src.Top - minT + PAGE_INSET
        dst.Width = src.Width
        dst.Height = src.Height

        CopyProps src, dst
        CopyList src, dst

        dsg.Controls.Remove src.Name
    Next i

    '-------------------------------------------------------------------
    ' 4. Re-seat the dialog-level controls under the tab strip.
    '-------------------------------------------------------------------
    footerTop = mp.Top + mp.Height + FOOTER_GAP
    formW = mp.Left + mp.Width + OUTER_MARGIN

    If HasControl(dsg, "lblStatus") Then
        With dsg.Controls("lblStatus")
            .Left = mp.Left + 2
            .Top = footerTop + 3
            .Width = formW - mp.Left - 200
        End With
    End If

    If HasControl(dsg, "btnCancel") Then
        With dsg.Controls("btnCancel")
            .Top = footerTop
            .Left = formW - OUTER_MARGIN - .Width
        End With
    End If

    If HasControl(dsg, "btnOK") Then
        With dsg.Controls("btnOK")
            .Top = footerTop
            If HasControl(dsg, "btnCancel") Then
                .Left = dsg.Controls("btnCancel").Left - .Width - 6
            Else
                .Left = formW - OUTER_MARGIN - .Width
            End If
        End With
    End If

    '-------------------------------------------------------------------
    ' 5. Fit the form around the new content.
    '    Width/Height include the window chrome, hence the fudge factors.
    '-------------------------------------------------------------------
    vbc.Properties("Width") = formW + 8
    vbc.Properties("Height") = footerTop + 24 + 26

    MsgBox "frmPatentToolsSettings restructured: " & n & _
           " control(s) moved onto tab 1." & vbCrLf & vbCrLf & _
           "Now run Debug > Compile Project and open the form to " & _
           "check alignment.", vbInformation, "Patent Tools"
    Exit Sub

Fail:
    MsgBox "Restructuring failed at runtime." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Restore your exported backup of the form before retrying. " & _
           "If the error mentions programmatic access, enable " & _
           """Trust access to the VBA project object model"".", _
           vbCritical, "Patent Tools"
End Sub

'-----------------------------------------------------------------------
' Helpers
'-----------------------------------------------------------------------

Private Function HasControl(ByVal container As Object, ByVal nm As String) As Boolean
    Dim c As Object

    On Error Resume Next
    Set c = container.Controls(nm)
    On Error GoTo 0

    HasControl = Not (c Is Nothing)
End Function

Private Function ProgIdOf(ByVal ctl As Object) As String
    Select Case TypeName(ctl)
        Case "Label":          ProgIdOf = "Forms.Label.1"
        Case "TextBox":        ProgIdOf = "Forms.TextBox.1"
        Case "ComboBox":       ProgIdOf = "Forms.ComboBox.1"
        Case "ListBox":        ProgIdOf = "Forms.ListBox.1"
        Case "CheckBox":       ProgIdOf = "Forms.CheckBox.1"
        Case "OptionButton":   ProgIdOf = "Forms.OptionButton.1"
        Case "ToggleButton":   ProgIdOf = "Forms.ToggleButton.1"
        Case "CommandButton":  ProgIdOf = "Forms.CommandButton.1"
        Case "Frame":          ProgIdOf = "Forms.Frame.1"
        Case "SpinButton":     ProgIdOf = "Forms.SpinButton.1"
        Case "Image":          ProgIdOf = "Forms.Image.1"
        Case Else
            Err.Raise 5, "ProgIdOf", "Unhandled control type """ & _
                      TypeName(ctl) & """. Move this control by hand."
    End Select
End Function

' Property-by-property copy. Each assignment is guarded because the set of
' valid properties differs per control type.
Private Sub CopyProps(ByVal src As Object, ByVal dst As Object)
    Dim props As Variant
    Dim i As Long

    props = Array("Caption", "Text", "Value", "ControlTipText", "Accelerator", _
                  "BackColor", "BackStyle", "ForeColor", "BorderColor", _
                  "BorderStyle", "SpecialEffect", "TextAlign", "WordWrap", _
                  "AutoSize", "Enabled", "Visible", "Locked", "TabStop", _
                  "TabIndex", "MultiLine", "EnterKeyBehavior", "ScrollBars", _
                  "PasswordChar", "MaxLength", "Style", "ListStyle", _
                  "ListRows", "ColumnCount", "BoundColumn", "TextColumn", _
                  "MatchRequired", "MatchEntry", "DropButtonStyle", _
                  "ShowDropButtonWhen", "TripleState", "Alignment", _
                  "GroupName", "Default", "Cancel", "Min", "Max", _
                  "SmallChange", "Orientation", "Picture", "PicturePosition", _
                  "MousePointer")

    ' Font first: some layout properties depend on the resolved font metrics.
    On Error Resume Next
    dst.Font.Name = src.Font.Name
    dst.Font.Size = src.Font.Size
    dst.Font.Bold = src.Font.Bold
    dst.Font.Italic = src.Font.Italic
    dst.Font.Underline = src.Font.Underline
    dst.Font.Weight = src.Font.Weight

    For i = LBound(props) To UBound(props)
        Err.Clear
        CallByName dst, CStr(props(i)), VbLet, CallByName(src, CStr(props(i)), VbGet)
    Next i
    On Error GoTo 0
End Sub

' ComboBox / ListBox contents, if the designer holds any design-time list.
Private Sub CopyList(ByVal src As Object, ByVal dst As Object)
    Dim i As Long

    On Error Resume Next
    If src.ListCount > 0 Then
        For i = 0 To src.ListCount - 1
            dst.AddItem src.List(i)
        Next i
    End If
    On Error GoTo 0
End Sub
