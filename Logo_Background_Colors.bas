Attribute VB_Name = "Logo_Background_Colors"
Option Explicit

Private Const Module_Name = "Logo_Background_Colors."

Private Const DarkestColor = &H763232 ' AF Dark Blue
Private Const LightestColor = &HE7E2E2 ' AF Light Gray

Public Sub DisableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = False
End Sub ' DisableButton

Public Sub EnableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = True
End Sub ' EnableButton

Public Sub HighLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = DarkestColor
    Btn.BackColor = LightestColor
    Btn.Enabled = True
End Sub ' HighLightButton

Public Sub HighLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = DarkestColor
    Ctl.BackColor = LightestColor
End Sub ' HighLightControl

Public Sub LowLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = LightestColor
    Btn.BackColor = DarkestColor
    Btn.Enabled = True
End Sub ' LowLightButton

Public Sub LowLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = LightestColor
    Ctl.BackColor = DarkestColor
End Sub ' LowLightControl

Public Sub Texture(ByRef Tbl As TableManager.TableClass)
    Const RoutineName = Module_Name & "Texture"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\texture.jpg") <> "" Then
        Set Tbl.Form.FormObj.Picture = LoadPicture(MainWorkbook.Path & "\texture.jpg")
    End If
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Public Function Logo( _
    ByRef Tbl As TableManager.TableClass, _
    ByRef LogoHeight As Single, _
    ByRef LogoWidth As Single _
    ) As Control
    
    Dim LogoImage As Control
    
    Const RoutineName = Module_Name & "Logo"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\logo.jpg") <> "" Then
        Set LogoImage = Tbl.Form.FormObj.Controls.Add("Forms.Image.1")
        Set LogoImage.Picture = LoadPicture(MainWorkbook.Path & "\logo.jpg")
        With LogoImage
            .PictureAlignment = fmPictureAlignmentTopRight
            .PictureSizeMode = fmPictureSizeModeZoom
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            .AutoSize = True
            LogoHeight = .Height
            LogoWidth = .Width
        End With
        Set Logo = LogoImage
    Else
        LogoHeight = 0
        Set Logo = Nothing
    End If
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function
