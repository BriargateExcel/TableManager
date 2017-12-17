Attribute VB_Name = "Logo_Background_Colors"
Option Explicit

Private Const Module_Name As String = "Logo_Background_Colors."

'TODO Use a parameter table to set the highlight and lowlight colors
'TODO Add ModuleName Debug check to these routines

Private Const DarkestColor As Long = &H763232    ' AF Dark Blue
Private Const LightestColor  As Long = &HE7E2E2  ' AF Light Gray

Public Sub DisableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = False
End Sub                                          ' DisableButton

Public Sub EnableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = True
End Sub                                          ' EnableButton

Public Sub HighLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = DarkestColor
    Btn.BackColor = LightestColor
    Btn.Enabled = True
End Sub                                          ' HighLightButton

Public Sub HighLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = DarkestColor
    Ctl.BackColor = LightestColor
End Sub                                          ' HighLightControl

Public Sub LowLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = LightestColor
    Btn.BackColor = DarkestColor
    Btn.Enabled = True
End Sub                                          ' LowLightButton

Public Sub LowLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = LightestColor
    Ctl.BackColor = DarkestColor
End Sub                                          ' LowLightControl

Public Sub Texture(ByRef Tbl As TableManager.TableClass)
    Const RoutineName As String = Module_Name & "Texture"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\texture.jpg") <> vbNullString Then
        Set Tbl.Form.FormObj.Picture = LoadPicture(MainWorkbook.Path & "\texture.jpg")
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub



