Attribute VB_Name = "HighlightColors"
Option Explicit

Private Const DarkestColor = &H763232 ' AF Dark Blue
Private Const LightestColor = &HE7E2E2 ' AF Light Gray

Public Sub DisableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = False
End Sub

Public Sub EnableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = True
End Sub

Public Sub HighLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = DarkestColor
    Btn.BackColor = LightestColor
    Btn.Enabled = True
End Sub

Public Sub HighLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = DarkestColor
    Ctl.BackColor = LightestColor
End Sub

Public Sub LowLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = LightestColor
    Btn.BackColor = DarkestColor
    Btn.Enabled = True
End Sub

Public Sub LowLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = LightestColor
    Ctl.BackColor = DarkestColor
End Sub
