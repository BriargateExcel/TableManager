Attribute VB_Name = "DataBaseRoutines"
'@Folder("TableManager.DataBase")

Option Explicit

Private Const Module_Name As String = "DataBaseRoutines."

Private pControls As ControlsClass
Private pEvents As EventsClass
Private pDataBaseFormName As String

Private Const pStandardGap As Long = 12
Private Const ButtonWidth As Long = 90
Private Const ButtonHeight As Long = 24

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.", "XLAM_Module.")
End Function                                     ' ModuleList

Public Function DataBaseFormName() As String
    DataBaseFormName = pDataBaseFormName
End Function

Public Sub BuildDataBaseForm( _
    ByVal Wkbk As Workbook, _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)

    Debug.Assert InScope(ModuleList, ModuleName)
    
    Const RoutineName As Variant = Module_Name & "BuildDataBaseForm"
    On Error GoTo ErrorHandler
    
    pDataBaseFormName = PlainDataBaseForm.Name
    PlainDataBaseForm.SetTable Tbl
    
    If LogoFileExists Then
        ' Create the UserForm
        Dim TempForm As VBComponent
        Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    
        Dim Frm As Object
        Set Frm = VBA.UserForms.Add(TempForm.Name)
        pDataBaseFormName = TempForm.Name
        Frm.Caption = "Save and Restore Table Data"
    
        Dim Evt As EventClass
        Set Evt = New EventClass
        Set Evt.FormObj = Frm
        Evt.Name = TempForm.Name
        
        Set pEvents = New EventsClass
        pEvents.Add Evt, Module_Name
    
        ' Add the texture
        Texture Frm
    
        Dim LogoImage As control
        Dim ControlsWidth As Long: ControlsWidth = 2 * ButtonWidth + StandardGap
        Dim Top As Long
    
        Top = StandardGap
    
        Set LogoImage = Logo(Frm)
        Frm.Width = Application.WorksheetFunction.Max( _
                    LogoImage.Width + 2 * StandardGap, _
                    ControlsWidth + 2 * StandardGap) + StandardGap
                
        LogoImage.Top = StandardGap
        LogoImage.Left = Frm.Width - 2 * StandardGap - LogoImage.Width
        Top = LogoImage.Height + 2 * StandardGap
        Dim Lft As Long
        Lft = ((Frm.Width - StandardGap) - ControlsWidth) / 2
    
        ' Build the label
        Dim Lbl As MSForms.Label
    
        BuildLabel Frm, Lbl
        Lbl.Top = Top
        Top = Top + StandardGap
        Lbl.Left = Lft
    
        Set pControls = New ControlsClass
    
        ' Build the text box
        Dim Ctl As MSForms.TextBox
    
        BuildTextBox Wkbk, Ctl, Frm
        Ctl.Top = Top
        Ctl.Left = Lft
        Ctl.Width = ControlsWidth
    
        Top = Top + 36 + StandardGap
    
        Dim ControlsHeight As Long
        ControlsHeight = Lbl.Height + Ctl.Height + 2 * ButtonHeight + 3 * StandardGap
    
        Frm.Height = LogoImage.Height + ControlsHeight + 4 * StandardGap
    
        BuildDataBaseFormButtons Frm, Lft, Top
    End If
    
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Private Sub BuildLabel( _
        ByVal Frm As Object, _
        ByRef Lbl As MSForms.Label)
    
    Set Lbl = Frm.Controls.Add("Forms.Label.1", "lblFileName", True)
    
    With Lbl
        .Caption = "File Name"
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        LowLightControl Lbl, Module_Name
        .Width = 42
    End With
End Sub

Private Sub BuildTextBox( _
    ByVal Wkbk As Workbook, _
        ByRef Ctl As MSForms.TextBox, _
        ByVal Frm As Object)
    
    Const RoutineName As String = Module_Name & "BuildTextBox"
    On Error GoTo ErrorHandler

    Set Ctl = Frm.Controls.Add("Forms.TextBox.1", "fldFileName", True)
    With Ctl
        .Height = 36
        .WordWrap = True
        .MultiLine = True
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .TextAlign = 1
        ' TODO Need to make the file name fetch dependent on the type of file storage selected
        ' Currently we only have CSV files
        ' Eventually, there could be other file types like MSAccess
        .Text = GetFullFileName(Wkbk, ActiveCellTableName)
    End With
    
    pControls.Add Ctl, Module_Name
    
    Dim Evt As EventClass
    Set Evt = New EventClass
    Set Evt.TextObj = Ctl
    Set Evt.FormObj = Frm
    Evt.Name = "FileName"
    pEvents.Add Evt, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Private Sub BuildDataBaseFormButtons( _
        ByVal Frm As Object, _
        ByVal Lft As Long, _
        ByVal Top As Long)
    
    Const RoutineName As String = Module_Name & "BuildDataBaseFormButtons"
    On Error GoTo ErrorHandler

    ' Top left
    BuildOneButton "Copy to File", Top, Lft, "Copy the contents of the table to an external data store", Frm
    
    ' Bottom left
    BuildOneButton "Change File", Top + ButtonHeight + StandardGap, Lft, "Change the source/destination file", Frm
    
    ' Top right
    BuildOneButton "Fetch From File", Top, ButtonWidth + StandardGap + Lft, "Fetch data from external file", Frm
    
    ' Bottom right
    BuildOneButton "Cancel File Processing", Top + ButtonHeight + StandardGap, ButtonWidth + StandardGap + Lft, "Cancel without doing anything", Frm

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Private Sub BuildOneButton( _
        ByVal Cption As String, _
        ByVal Top As Single, _
        ByVal Lft As Single, _
        ByVal Tip As String, _
        ByVal Frm As Object)
    
    Dim Ctl As MSForms.CommandButton
    
    Const RoutineName As String = Module_Name & "BuildOneButton"
    On Error GoTo ErrorHandler

    Set Ctl = Frm.Controls.Add("Forms.CommandButton.1")
    With Ctl
        .Caption = Cption
        .Top = Top
        .Left = Lft
        .Height = ButtonHeight
        .Width = ButtonWidth
        LowLightButton Ctl, Module_Name
        .ControlTipText = Tip
    End With
    
    pControls.Add Ctl, Module_Name
    
    Dim Evt As EventClass
    
    Set Evt = New EventClass
    Set Evt.ButtonObj = Ctl
    Set Evt.FormObj = Frm
    Evt.Name = Cption
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildOneButton

Public Function StandardGap() As Long
    StandardGap = pStandardGap
End Function


