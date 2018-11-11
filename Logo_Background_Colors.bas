Attribute VB_Name = "Logo_Background_Colors"
'@Folder("TableManager.Colors")

Option Explicit

Private Const Module_Name As String = "Logo_Background_Colors."

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.", "FormClass.", "FormRoutines.", "DataBaseRoutines.")
End Function                                     ' ModuleList

Public Sub DisableButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "DisableButton"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Btn.Enabled = False

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' DisableButton

Public Sub EnableButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "EnableButton"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' EnableButton

Public Sub HighLightButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "HighLightButton"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Btn.ForeColor = DarkestColorValue
    Btn.BackColor = LightestColorValue
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' HighLightButton

Public Sub HighLightControl( _
       ByVal Ctl As control, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "HighLightControl"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Ctl.ForeColor = DarkestColorValue
    Ctl.BackColor = LightestColorValue

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' HighLightControl

Public Sub LowLightButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "LowLightButton"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Btn.ForeColor = LightestColorValue
    Btn.BackColor = DarkestColorValue
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' LowLightButton

Public Sub LowLightControl( _
       ByVal Ctl As control, _
       ByVal ModuleName As String)
    
    If Ctl Is Nothing Then Exit Sub
    
    Const RoutineName As String = Module_Name & "LowLightControl"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Ctl.ForeColor = LightestColorValue
    Ctl.BackColor = DarkestColorValue

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' LowLightControl

Public Function TextureFileExists() As Boolean
    TextureFileExists = (Dir(GetMainWorkbook.Path & "\texture.jpg") <> vbNullString)
End Function

Public Function LogoFileExists() As Boolean
    LogoFileExists = (Dir(GetMainWorkbook.Path & "\logo.jpg") <> vbNullString)
End Function

Public Sub Texture(ByRef Frm As Object)
    Const RoutineName As String = Module_Name & "Texture"
    On Error GoTo ErrorHandler
    
    If TextureFileExists Then
        Set Frm.Picture = LoadPicture(GetMainWorkbook.Path & "\texture.jpg")
    Else
        Frm.BackColor = DarkestColorValue
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Public Function Logo( _
       ByRef Frm As Object) As control
    
    Const RoutineName As String = Module_Name & "Logo"
    On Error GoTo ErrorHandler
    
    Dim LogoImage As control
    Dim Pic As StdPicture
    
    If LogoFileExists Then
        Set LogoImage = Frm.Controls.Add("Forms.Image.1")
        Set Pic = LoadPicture(GetMainWorkbook.Path & "\logo.jpg")
        Set LogoImage.Picture = Pic
        With LogoImage
            .PictureAlignment = fmPictureAlignmentTopLeft
            .PictureSizeMode = fmPictureSizeModeZoom
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            
            Dim PicHeightToWidth As Single
            PicHeightToWidth = Pic.Height / Pic.Width
            
            Dim Factor As Single
            Factor = Application.WorksheetFunction.Max(Pic.Height, Pic.Width) / 35
            Factor = Application.WorksheetFunction.Min(Factor, 200)

            If PicHeightToWidth > 1 Then
                ' Height > Width
                .Height = Factor
                .Width = Factor / PicHeightToWidth
            Else
                ' Width > Height
                .Width = Factor
                .Height = Factor * PicHeightToWidth
            End If
        End With
        Set Logo = LogoImage
    Else
        Set Logo = Nothing
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function


