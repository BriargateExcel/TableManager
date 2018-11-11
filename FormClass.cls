VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Forms")

Option Explicit

Private Const Module_Name As String = "FormClass."

Private pFormName As String
Private pWorksheetName As String
Private pFormObj As Object
' pFormObj:
' When declared as MSForms.UserForm, it has no height or width
' When declared as an object, it has both height and width
Private pUserForm As MSForms.UserForm
Private pTable As TableClass
Private pControls As ControlsClass
Private pEvents As EventsClass

Private pEnabled As Boolean                      ' Whether form is enabled

' TODO Calculate the gaps and multipliers based on font and font size
Private Const ThinGap As Long = 3
Private Const TextHeight As Long = 18
Private Const ComboHeight As Long = 18
Private Const ButtonHeight As Long = 24
Private Const ButtonWidth As Long = 72
Private Const FontSize As Single = 8.25
Private Const WidthMultiplier As Long = 6
Private Const TextMultiplier As Long = 5
Private Const ScrollBarWidth As Long = 24.75

Private Function ModuleList() As Variant
    ModuleList = Array("TableRoutines.")
End Function                                     ' ModuleList

Public Function BuildForm( _
    ByVal Wkbk As Workbook, _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String _
       ) As MSForms.UserForm
       
    ' Called in TableRoutines
       
    Debug.Assert Initializing

    Const RoutineName As Variant = Module_Name & "BuildForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Set pTable = Tbl
    pWorksheetName = Tbl.WorksheetName
    
    ' Create the UserForm
    Dim TempForm As VBComponent
    Set TempForm = Wkbk.VBProject.VBComponents.Add(vbext_ct_MSForm)
    pFormName = TempForm.Name
    
    Set Tbl.Form = New FormClass
    Tbl.Form.Name = pFormName
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    Set pFormObj = Wkbk.VBProject.VBComponents.Add(vbext_ct_MSForm)
'    Set pUserForm = Wkbk.VBProject.VBComponents.Add(vbext_ct_MSForm)
'    pFormObj.Name = TempForm.Name
'    Set Tbl.Form.FormObj = pFormObj
'    Set Tbl.Form.UserForm = pFormObj
'    Set pUserForm = pFormObj
'    pFormObj.Caption = Tbl.Name
    
    Dim Evt As EventClass
    Set Evt = New EventClass
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = pFormName
    pEvents.Add Evt, Module_Name
    
    ' Add the texture
    Texture Tbl.Form.FormObj
    
    ' Add the logo
    Dim LogoImage As control
    Dim Top As Single
    Set LogoImage = Logo(Tbl.Form.FormObj)
    Dim LogoHeight As Single
    Dim LogoWidth As Single
    If LogoImage Is Nothing Then
        LogoWidth = 0
        LogoHeight = 0
    Else
        LogoWidth = LogoImage.Width
        LogoHeight = LogoImage.Height
        LogoImage.Top = StandardGap
    End If
    Top = LogoHeight + 2 * StandardGap
    
    ' Build the Controls
    Dim MaxLabelWidth As Single
    Dim MaxFieldWidth As Single
    BuildControls Tbl, Top, MaxLabelWidth, MaxFieldWidth
    
    ' Build the Buttons
    Dim NumRows As Long
    Dim NumCols As Long
    Dim ButtonSetWidth As Single
    ButtonSetWidth = BuildAllButtons(Top, MaxLabelWidth, MaxFieldWidth, Tbl, NumRows, NumCols) + 2 * StandardGap
        
    ' Set Properties for TempForm now that we know all the heights and widths
    Dim FieldWidth As Single
    FieldWidth = MaxLabelWidth + MaxFieldWidth + 3 * StandardGap
    Dim LogoTotalWidth As Single
    LogoTotalWidth = LogoWidth + 2 * StandardGap
    pFormObj.Width = Application.WorksheetFunction.Max( _
                     FieldWidth, ButtonSetWidth, LogoTotalWidth) + ScrollBarWidth + 2 * StandardGap
        
    Top = Top + ButtonHeight + StandardGap
    pFormObj.Height = Top + NumRows * StandardGap
    
    ' Right justify the Logo
    If Not LogoImage Is Nothing Then
        LogoImage.Left = pFormObj.Width - StandardGap - ScrollBarWidth - LogoWidth
    End If
    
    ' Center the fields
    Dim LeftFieldMargin As Single
    LeftFieldMargin = (pFormObj.Width - ScrollBarWidth - FieldWidth) / 2
    Dim LeftButtonMargin As Single
    LeftButtonMargin = ((pFormObj.Width - ScrollBarWidth) - ButtonSetWidth) / 2 ' Centers the buttons
    Dim Ctl As Variant
    
    For Each Ctl In pFormObj.Controls
        If Left$(Ctl.Name, 3) = "lbl" Or _
           Left$(Ctl.Name, 3) = "fld" Or _
           Left$(Ctl.Name, 3) = "dat" Or _
           Left$(Ctl.Name, 3) = "cmb" Or _
           Left$(Ctl.Name, 3) = "whl" Or _
           Left$(Ctl.Name, 3) = "val" Then
            Ctl.Left = Ctl.Left + LeftFieldMargin
        End If
        
        If Left$(Ctl.Name, 7) = "Command" Then
            Ctl.Left = Ctl.Left + LeftButtonMargin
        End If
                
    Next Ctl
    
    pFormObj.ScrollBars = fmScrollBarsVertical
        
    Set BuildForm = pFormObj
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' BuildForm

Private Sub BuildControls( _
        ByVal Tbl As TableClass, _
        ByRef Top As Single, _
        ByRef MaxLabelWidth As Single, _
        ByRef MaxFieldWidth As Single)
        
    ' Called in FormClass

    Dim I As Long
    Dim Field As Variant
    Dim Lbl As control
    
    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
            
        ' Build the field label
        BuildLabel Top, Field, MaxLabelWidth, Lbl
            
        ' Build the data field
        ' TODO Set the format of the text field to match the format of the table element
        '   Example: Locked numerical cells take on a General format
        If Field.Locked Then
            ' Set this up as a constant (label)
            BuildDataLabel Top, Field, MaxFieldWidth, Lbl
                
        Else
            ' Set this up as an input control
            BuildUnLockedFields Field, Top, MaxFieldWidth, Lbl, Tbl
        End If
            
    Next I

    ' Position the controls now that we know all the widths
    Dim DataLeft As Single
    DataLeft = MaxLabelWidth + 2 * StandardGap
    
    Dim Ctl As control
    For Each Ctl In pFormObj.Controls
        If Ctl.Name <> "Filler" And Left$(Ctl.Name, 5) <> "Image" Then
            Ctl.Width = MaxFieldWidth
            Ctl.Left = DataLeft
        End If
        If Left$(Ctl.Name, 3) = "lbl" Then
            Ctl.Width = MaxLabelWidth
            Ctl.Left = StandardGap
        End If
    Next Ctl
        
    Top = Top + StandardGap
End Sub

Private Sub BuildUnLockedFields( _
        ByVal Field As Variant, _
        ByRef Top As Single, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)
        
    ' Called in FormClass
    
    Select Case Field.CellType
    Case xlValidateInputOnly
        ' Set this up as a text field
        BuildText Top, Field, MaxFieldWidth, Lbl, Tbl
                    
    Case xlValidateList
        ' Set this up as a combobox
        BuildCombo Top, Field, MaxFieldWidth, Lbl, Tbl
                    
    Case xlValidateCustom
        ' Set this up as a text field
        BuildCustom Top, Field, MaxFieldWidth, Lbl, Tbl
                
    Case xlValidateDate
        ' Set this up as a textbox
        BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

    Case xlValidateDecimal
        ' Set this up as a textbox
        BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

    Case xlValidateTextLength
        ' Set this up as a text field
        BuildText Top, Field, MaxFieldWidth, Lbl, Tbl

    Case xlValidateTime
        ' Set this up as a text field
        BuildText Top, Field, MaxFieldWidth, Lbl, Tbl

    Case xlValidateWholeNumber:
        ' Set this up as a textbox
        BuildWholeNumber Top, Field, MaxFieldWidth, Lbl, Tbl
                
        '                Case Else: Stop
        'TODO Add a hypertext "data type" as an option
        '   Not clear how to set this up for the user because
        '   I'm not sure if I can modify the Data Validation
        '   drop down options
    End Select

End Sub

Private Function BuildAllButtons( _
        ByRef Top As Single, _
        ByVal MaxLabelWidth As Single, _
        ByVal MaxFieldWidth As Single, _
        ByVal Tbl As TableClass, _
        ByRef NumRows As Long, _
        ByRef NumCols As Long _
        ) As Single
        
    ' Called in FormClass
    
    '    Dim DataLeft As Single
    Dim Arrangement As String
    
    Const RoutineName As String = Module_Name & "BuildAllButtons"
    On Error GoTo ErrorHandler

    Debug.Assert Initializing
    
    NumRows = 4
    NumCols = 3
        
    Debug.Assert NumRows = 2 Or NumRows = 3 Or NumRows = 4
    Debug.Assert NumCols = 3 Or NumCols = 4 Or NumCols = 6
    Debug.Assert NumRows * NumCols = 12
    
    If MaxLabelWidth + MaxFieldWidth > 4 * ButtonWidth + (4 + 2) * StandardGap Then
        NumRows = 3
        NumCols = 4
    End If
    
    If MaxLabelWidth + MaxFieldWidth > 6 * ButtonWidth + (6 + 2) * StandardGap Then
        NumRows = 2
        NumCols = 6
    End If
    
    BuildAllButtons = NumCols * ButtonWidth + (NumCols - 1) * StandardGap
    
    Arrangement = Format$(NumRows, "0") & " Rows of " & Format$(NumCols, "0")
    
    ' TODO Can I rearrange the buttons so they more logically fall into
    '   2x6, 3x4, and 4x3 arrangements? In other words, is there a better,
    '   more logical way to arrange the buttons for the 3 configurations?
    Select Case Arrangement
    Case "2 Rows of 6"
        ' Add First Row of Buttons
        ' First row of buttons
        BuildButtonRow 0, 5, Top, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 6, 11, Top, Tbl

    Case "3 Rows of 4"
        ' Add First Row of Buttons
        ' First row of buttons
        BuildButtonRow 0, 3, Top, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 4, 7, Top, Tbl

        ' Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 8, 11, Top, Tbl
        
    Case "4 Rows of 3"
        ' Add First Row of Buttons
        ' First row of buttons
        BuildButtonRow 0, 2, Top, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 3, 5, Top, Tbl

        ' Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 6, 8, Top, Tbl

        ' Fourth row of buttons
        Top = Top + ButtonHeight + StandardGap
        BuildButtonRow 9, 11, Top, Tbl
    
        '    Case Else: Stop
    End Select

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' BuildAllButtons

Private Sub BuildButtonRow( _
        ByVal LowerBound As Long, _
        ByVal UpperBound As Long, _
        ByVal Top As Single, _
        ByVal Tbl As TableClass)
        
    ' Called in FormClass

    Const RoutineName As String = Module_Name & "BuildButtonRow"
    On Error GoTo ErrorHandler
    
    Dim I As Long

    Dim ButtonNames() As Variant
    Dim ButtonLabels() As Variant
    
    ButtonNames = Array( _
                  "Validate Form", "Copy To Table", "Next Row", "Previous Row", _
                  "Insert Above", "Insert Below", "Insert at Top", "Insert at Bottom", _
                  "Delete This Row", "Go To Top", "Go To Bottom", "Cancel")
    
    ButtonLabels = Array( _
                   "Validate the contents of the Form", _
                   "Copy the contents of the Form to the Table row", _
                   "Move to previous row of the Table", _
                   "Move to next row of the Table", _
                   "Insert a blank line above this Table row and move to it", _
                   "Insert a blank line below this Table row and move to it", _
                   "Insert a blank line above the first row of the Table and move to it", _
                   "Insert a blank line below the last row of the Table and move to it", _
                   "Delete this row of the Table", _
                   "Move to the first row of the Table", _
                   "Move to the last row of the Table", _
                   "Exit the Form without changing the Table any further")
    
    Dim DataLeft As Single
    DataLeft = 0
    For I = LowerBound To UpperBound
        BuildOneButton _
        ButtonNames(I), Top, DataLeft, ButtonLabels(I), Tbl
            
        DataLeft = DataLeft + ButtonWidth + StandardGap
    Next I
    
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
        ByVal Tbl As TableClass)
    
    Dim Ctl As control
    Dim Evt As EventClass
    
    Const RoutineName As String = Module_Name & "BuildOneButton"
    On Error GoTo ErrorHandler

    Debug.Assert Initializing

    Set Ctl = pFormObj.Controls.Add("Forms.CommandButton.1")
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
    
    Set Evt = New EventClass
    Set Evt.ButtonObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Cption
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildOneButton

Private Function BuildCaption( _
        ByVal InTitle As String, _
        ByVal InMessage As String _
        ) As String
    
    Dim Msg As String
    
    Msg = InTitle
    
    If InTitle <> vbNullString Then Msg = Msg & ": "
    
    BuildCaption = Msg & InMessage
    
End Function

Private Sub BuildCombo( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)

    Const RoutineName As String = Module_Name & "BuildCombo"
    On Error GoTo ErrorHandler
    
    Dim Ctl As control
    Dim Evt As EventClass

    Dim ValListItem As Variant

    Set Ctl = pFormObj.Controls.Add("Forms.ComboBox.1", "cmb" & Field.Name, True)
    With Ctl
        If VarType(Field.ValidationList) = 8 Then
            ' If there is only one entry in the ValidaionList the For Each fails
            ' VarType = 8 is a String
            .AddItem CStr(Field.ValidationList)
        Else
            For Each ValListItem In Field.ValidationList
                .AddItem CStr(ValListItem)
            Next ValListItem
        End If
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .Top = Top
        Top = Top + ComboHeight
        .Height = ComboHeight
        Lbl.Height = ComboHeight
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
    End With
    Set Field.FormControl = Ctl
    pControls.Add Ctl, Module_Name
    Set Evt = New EventClass
    Set Evt.ComboObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

Private Sub BuildText( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)

    Const RoutineName As String = Module_Name & "BuildText"
    On Error GoTo ErrorHandler
    
    Dim Ctl As MSForms.TextBox
    Dim Evt As EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "fld" & Field.Name, True)
    With Ctl
        .Top = Top
        If Field.Wrap Then
            .Height = 3 * TextHeight
            Lbl.Height = 3 * TextHeight
            Top = Top + 3 * TextHeight
            .WordWrap = True
            .MultiLine = True
            .ScrollBars = fmScrollBarsVertical
        Else
            .Height = TextHeight
            Lbl.Height = TextHeight
            Top = Top + TextHeight
            .WordWrap = False
            .MultiLine = False
            .ScrollBars = fmScrollBarsNone
        End If
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
        .TextAlign = 1
    End With
    
    Set Field.FormControl = Ctl
    pControls.Add Ctl, Module_Name
    
    Set Evt = New EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' BuildText

Private Sub BuildCustom( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)

    Const RoutineName As String = Module_Name & "BuildCustom"
    On Error GoTo ErrorHandler
    
    Dim Ctl As control
    Dim Evt As EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "fld" & Field.Name, True)
    With Ctl
        .Top = Top
        If Field.Wrap Then
            .Height = 3 * TextHeight
            Lbl.Height = 3 * TextHeight
            Top = Top + 3 * TextHeight
            .WordWrap = True
            .MultiLine = True
            .ScrollBars = fmScrollBarsVertical
        Else
            .Height = TextHeight
            Lbl.Height = TextHeight
            Top = Top + TextHeight
            .WordWrap = False
            .MultiLine = False
            .ScrollBars = fmScrollBarsNone
        End If
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
        .TextAlign = 1
    End With
    Set Field.FormControl = Ctl
    pControls.Add Ctl, Module_Name
    Set Evt = New EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildCustom

Private Sub BuildDataLabel( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control)

    Const RoutineName As String = Module_Name & "BuildDataLabel"
    On Error GoTo ErrorHandler
    
    Dim Ctl As control

    Set Ctl = pFormObj.Controls.Add("Forms.Label.1", "val" & Field.Name, True)
    With Ctl
        .Top = Top
        .Height = TextHeight + ThinGap
        Lbl.Height = TextHeight + ThinGap
        Top = Top + TextHeight + ThinGap
        .Font.Size = FontSize
        .WordWrap = False
        LowLightControl Ctl, Module_Name
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
    End With
    Set Field.FormControl = Ctl
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildDataLabel

Private Sub BuildDate( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)

    Const RoutineName As String = Module_Name & "BuildDate"
    On Error GoTo ErrorHandler
    
    Dim Ctl As control
    Dim Evt As EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "dat" & Field.Name, True)
    With Ctl
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .Top = Top
        Top = Top + ComboHeight
        .Height = ComboHeight
        Lbl.Height = ComboHeight
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
    End With
    Set Field.FormControl = Ctl
    pControls.Add Ctl, Module_Name
    Set Evt = New EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildDate

Private Sub BuildLabel( _
        ByVal Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxLabelWidth As Single, _
        ByRef Lbl As control)

    Const RoutineName As String = Module_Name & "BuildLabel"
    On Error GoTo ErrorHandler

    Set Lbl = pFormObj.Controls.Add("Forms.Label.1", "lbl" & Field.Name, True)
    
    With Lbl
        .Top = Top
        .Left = StandardGap
        .Caption = Field.HeaderText
        .TextAlign = fmTextAlignRight
        .WordWrap = False
        LowLightControl Lbl, Module_Name
        .Width = Len(Field.HeaderText) * TextMultiplier
        If .Width > MaxLabelWidth Then MaxLabelWidth = .Width
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildLabel

Private Sub BuildWholeNumber( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As control, _
        ByRef Tbl As TableClass)

    Const RoutineName As String = Module_Name & "BuildWholeNumber"
    On Error GoTo ErrorHandler
    
    Dim Ctl As control
    Dim Evt As EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "whl" & Field.Name, True)
    With Ctl
        .BackColor = DarkestColorValue
        .ForeColor = LightestColorValue
        .Top = Top
        .Width = Field.ColumnWidth * WidthMultiplier
        Top = Top + ComboHeight
        .Height = ComboHeight
        Lbl.Height = ComboHeight
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
    End With
    Set Field.FormControl = Ctl
    pControls.Add Ctl, Module_Name
    Set Evt = New EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildWholeNumber

Public Property Get Controls() As ControlsClass: Set Controls = pControls: End Property

Public Property Get Enabled() As Boolean: Enabled = pEnabled: End Property

Public Property Let Enabled(ByVal Vbl As Boolean): pEnabled = Vbl: End Property

Public Property Get UserForm() As MSForms.UserForm: Set UserForm = pUserForm: End Property

Public Property Set UserForm(ByVal Vbl As MSForms.UserForm): Set pUserForm = Vbl: End Property

Public Property Get FormObj() As Object: Set FormObj = pFormObj: End Property

Public Property Set FormObj(ByVal Vbl As Object): Set pFormObj = Vbl: End Property

Public Property Get Name() As String: Name = pFormName: End Property

Public Property Let Name(ByVal Vbl As String): pFormName = Vbl: End Property

Public Property Get Table() As TableClass: Set Table = pTable: End Property

Public Property Get WorksheetName() As String: WorksheetName = pWorksheetName: End Property

Public Property Set Picture(ByVal Vbl As StdPicture): Set pFormObj.Picture = Vbl: End Property

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pControls = New ControlsClass
    Set pEvents = New EventsClass
End Sub                                          ' Class_Initialize


