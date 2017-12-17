VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name As String = "FormClass."

Private pFormName As String
Private pWorksheetName As String
Private pFormObj As Object                       'MSForms.UserForm; when declared as UserFrom it has no height or width
' When declared as an object, it has both height and width
Private pTable As TableManager.TableClass
Private pControls As TableManager.ControlsClass
Private pEvents As TableManager.EventsClass

Private pEnabled As Boolean                      ' Whether form is enabled

' TODO Calculate the gaps and multipliers based on font and font size
Private Const StandardGap As Long = 12
Private Const ThinGap As Long = 3
Private Const TextHeight As Long = 18
Private Const ComboHeight As Long = 18
Private Const ButtonHeight As Long = 24
Private Const ButtonWidth As Long = 72
Private Const FontSize As Single = 8.25
Private Const WidthMultiplier As Long = 6
Private Const TextMultiplier As Long = 5

Public Function BuildForm( _
       ByVal Tbl As TableManager.TableClass _
       ) As MSForms.UserForm
       
    Dim TempForm As VBComponent
    Debug.Assert Initializing

    Const RoutineName As Variant = Module_Name & "BuildForm"
    On Error GoTo ErrorHandler
    
    Set pTable = Tbl
    pWorksheetName = Tbl.WorksheetName
    
    ' Create the UserForm
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    pFormName = TempForm.Name
    
    Set Tbl.Form = New TableManager.FormClass
    Tbl.Form.Name = pFormName
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    Set Tbl.Form.FormObj = pFormObj
    pFormObj.Caption = Tbl.Name
    
    Dim Evt As TableManager.EventClass
    Set Evt = New TableManager.EventClass
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = pFormName
    pEvents.Add Evt
    
    ' Add the texture
    TableManager.Texture Tbl
    
    ' Add the logo
    Dim LogoHeight As Single
    Dim Top As Single
    BuildLogo LogoHeight, Tbl
    Top = LogoHeight + 2 * StandardGap
    
    ' Build the Controls
    Dim MaxLabelWidth As Single
    Dim MaxFieldWidth As Single
    Dim MaxButtonWidth As Single
    BuildControls Tbl, Top, MaxLabelWidth, MaxFieldWidth
    
    ' Build the Buttons
    Dim NumRows As Long
    Dim NumCols As Long
    BuildAllButtons Top, MaxLabelWidth, MaxFieldWidth, MaxButtonWidth, Tbl, NumRows, NumCols
        
    ' Set Properties for TempForm now that we know all the heights and widths
    pFormObj.Width = Application.WorksheetFunction.Max( _
                     MaxLabelWidth + MaxFieldWidth + NumCols * StandardGap, _
                     MaxButtonWidth)
        
    Top = Top + ButtonHeight + StandardGap
    pFormObj.Height = Top + NumRows * StandardGap
        
    Set BuildForm = pFormObj
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' BuildForm

Private Sub BuildControls( _
        ByVal Tbl As TableManager.TableClass, _
        ByRef Top As Single, _
        ByRef MaxLabelWidth As Single, _
        ByRef MaxFieldWidth As Single)

    Dim I As Long
    Dim Field As Variant
    Dim Lbl As Control
    
    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I)
            
        ' Build the field label
        BuildLabel Top, Field, MaxLabelWidth, Lbl
            
        ' Build the data field
        If Field.Locked Then
            ' Set this up as a constant (label)
            BuildDataLabel Top, Field, MaxFieldWidth, Lbl
                
        Else
            ' Field is unlocked
            Select Case Field.CellType
            Case xlValidateInputOnly
                ' Set this up as a text field
                BuildText Top, Field, MaxFieldWidth, Lbl, Tbl
                    
            Case xlValidateList
                ' Set this up as a combobox
                BuildCombo Top, Field, MaxFieldWidth, Lbl, Tbl
                    
            Case xlValidateCustom
                MsgBox "Custom validation is not yet implemented", _
                       vbOKOnly Or vbExclamation, "Custom Validation"
                'TODO Implement Custom validation
                
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
                MsgBox "Time validation is not yet implemented", _
                       vbOKOnly Or vbExclamation, "Time Validation"
                'TODO Implement Time validation

            Case xlValidateWholeNumber:
                ' Set this up as a textbox
                BuildWholeNumber Top, Field, MaxFieldWidth, Lbl, Tbl
                
                '                Case Else: Stop
                'TODO Add a hypertext "data type" as an option
                '   Not clear how to set this up for the user because
                '   I'm not sure if I can modify the Data Validation
                '   drop down options
            End Select
        End If                                   ' Field is locked
            
    Next I

    ' Position the controls now that we know all the widths
    Dim DataLeft As Single
    DataLeft = MaxLabelWidth + 2 * StandardGap
    
    Dim Ctl As Control
    For Each Ctl In pFormObj.Controls
        If Left$(Ctl.Name, 3) = "lbl" Then
            Ctl.Width = MaxLabelWidth
            Ctl.Left = StandardGap
        Else
            If Ctl.Name <> "Filler" Then
                Ctl.Width = MaxFieldWidth
                Ctl.Left = DataLeft
            End If
        End If
    Next Ctl
        
    Top = Top + StandardGap
End Sub

Private Sub BuildLogo( _
        ByRef LogoHeight As Single, _
        ByRef Tbl As TableManager.TableClass)
    
    Dim LogoImage As Control
    Dim LogoWidth As Single
    Set LogoImage = Logo(Tbl, LogoHeight, LogoWidth)
    If Not LogoImage Is Nothing Then
        LogoImage.Left = pFormObj.Width - LogoWidth
        LogoImage.Left = 0
        LogoImage.Left = pFormObj.InsideWidth - LogoImage.Width - StandardGap
        LogoImage.Top = StandardGap
    End If
End Sub

Public Function Logo( _
       ByRef Tbl As TableManager.TableClass, _
       ByRef LogoHeight As Single, _
       ByRef LogoWidth As Single _
       ) As Control
    
    Dim LogoImage As Control
    
    Const RoutineName As String = Module_Name & "Logo"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\logo.jpg") <> vbNullString Then
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
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Private Sub BuildAllButtons( _
        ByRef Top As Single, _
        ByVal MaxLabelWidth As Single, _
        ByVal MaxFieldWidth As Single, _
        ByRef MaxButtonWidth As Single, _
        ByVal Tbl As TableManager.TableClass, _
        ByRef NumRows As Long, _
        ByRef NumCols As Long)
    
    Dim DataLeft As Single
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
    
    MaxButtonWidth = NumCols * ButtonWidth + (NumCols + 2) * StandardGap
    
    Arrangement = Format$(NumRows, "0") & " Rows of " & Format$(NumCols, "0")
    
    Select Case Arrangement
    Case "2 Rows of 6"
        ' Add First Row of Buttons
        DataLeft = StandardGap
        ' First row of buttons
        BuildButtonRow 0, 5, Top, DataLeft, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 6, 11, Top, DataLeft, Tbl

    Case "3 Rows of 4"
        ' Add First Row of Buttons
        DataLeft = StandardGap
        ' First row of buttons
        BuildButtonRow 0, 3, Top, DataLeft, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 4, 7, Top, DataLeft, Tbl

        ' Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 8, 11, Top, DataLeft, Tbl
        
    Case "4 Rows of 3"
        ' Add First Row of Buttons
        DataLeft = StandardGap
        ' First row of buttons
        BuildButtonRow 0, 2, Top, DataLeft, Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 3, 5, Top, DataLeft, Tbl

        ' Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 6, 8, Top, DataLeft, Tbl

        ' Fourth row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButtonRow 9, 11, Top, DataLeft, Tbl
    
        '    Case Else: Stop
    End Select

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' BuildAllButtons

Private Sub BuildButtonRow( _
        ByVal LowerBound As Long, _
        ByVal UpperBound As Long, _
        ByRef Top As Single, _
        ByRef DataLeft As Single, _
        ByVal Tbl As TableManager.TableClass)

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
        
    For I = LowerBound To UpperBound
        BuildOneButton ButtonNames(I), Top, DataLeft, _
                       ButtonLabels(I), Tbl
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
        ByRef Lft As Single, _
        ByVal Tip As String, _
        ByVal Tbl As TableManager.TableClass)
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass
    
    Const RoutineName As String = Module_Name & "BuildOneButton"
    On Error GoTo ErrorHandler

    Debug.Assert Initializing

    Set Ctl = pFormObj.Controls.Add("Forms.CommandButton.1")
    With Ctl
        .Caption = Cption
        .Top = Top
        .Left = Lft
        Lft = Lft + ButtonWidth + StandardGap
        .Height = ButtonHeight
        .Width = ButtonWidth
        LowLightButton Ctl
        .ControlTipText = Tip
    End With
    
    pControls.Add Ctl
    
    Set Evt = New TableManager.EventClass
    Set Evt.ButtonObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Cption
    pEvents.Add Evt
    
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
        ByRef Lbl As Control, _
        ByRef Tbl As TableManager.TableClass)

    Const RoutineName As String = Module_Name & "BuildCombo"
    On Error GoTo ErrorHandler
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass

    Dim ValListItem As Variant

    Set Ctl = pFormObj.Controls.Add("Forms.ComboBox.1", "cmb" & Field.Name, True)
    With Ctl
        For Each ValListItem In Field.ValidationList
            .AddItem CStr(ValListItem)
        Next ValListItem
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
    pControls.Add Ctl
    Set Evt = New TableManager.EventClass
    Set Evt.ComboObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt
    
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
        ByRef Lbl As Control, _
        ByRef Tbl As TableManager.TableClass)

    Const RoutineName As String = Module_Name & "BuildText"
    On Error GoTo ErrorHandler
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass

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
        .Width = Field.ColumnWidth * WidthMultiplier
        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
        .Font.Size = FontSize
        .ControlTipText = BuildCaption(Field.InputTitle, Field.InputMessage)
        .TextAlign = 1
    End With
    Set Field.FormControl = Ctl
    pControls.Add Ctl
    Set Evt = New TableManager.EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildText

Private Sub BuildDataLabel( _
        ByRef Top As Single, _
        ByVal Field As Variant, _
        ByRef MaxFieldWidth As Single, _
        ByRef Lbl As Control)

    Const RoutineName As String = Module_Name & "BuildDataLabel"
    On Error GoTo ErrorHandler
    
    Dim Ctl As Control

    Set Ctl = pFormObj.Controls.Add("Forms.Label.1", "val" & Field.Name, True)
    With Ctl
        .Top = Top
        .Height = TextHeight + ThinGap
        Lbl.Height = TextHeight + ThinGap
        Top = Top + TextHeight + ThinGap
        .Font.Size = FontSize
        .WordWrap = False
        LowLightControl Ctl
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
        ByRef Lbl As Control, _
        ByRef Tbl As TableManager.TableClass)

    Const RoutineName As String = Module_Name & "BuildDate"
    On Error GoTo ErrorHandler
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "dat" & Field.Name, True)
    With Ctl
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
    pControls.Add Ctl
    Set Evt = New TableManager.EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt
    
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
        ByRef Lbl As Control)

    Const RoutineName As String = Module_Name & "BuildLabel"
    On Error GoTo ErrorHandler

    Set Lbl = pFormObj.Controls.Add("Forms.Label.1", "lbl" & Field.Name, True)
    
    With Lbl
        .Top = Top
        .Left = StandardGap
        .Caption = Field.HeaderText
        .TextAlign = fmTextAlignRight
        .WordWrap = False
        LowLightControl Lbl
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
        ByRef Lbl As Control, _
        ByRef Tbl As TableManager.TableClass)

    Const RoutineName As String = Module_Name & "BuildWholeNumber"
    On Error GoTo ErrorHandler
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass

    Set Ctl = pFormObj.Controls.Add("Forms.TextBox.1", "whl" & Field.Name, True)
    With Ctl
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
    pControls.Add Ctl
    Set Evt = New TableManager.EventClass
    Set Evt.TextObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Field.Name
    pEvents.Add Evt
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildWholeNumber

Public Property Get Controls() As TableManager.ControlsClass: Set Controls = pControls: End Property

Public Property Get Enabled() As Boolean: Enabled = pEnabled: End Property

Public Property Let Enabled(ByVal En As Boolean): pEnabled = En: End Property

Public Property Get FormObj() As MSForms.UserForm: Set FormObj = pFormObj: End Property

Public Property Set FormObj(ByVal FO As MSForms.UserForm): Set pFormObj = FO: End Property

Public Property Get Name() As String: Name = pFormName: End Property

Public Property Let Name(ByVal FN As String): pFormName = FN: End Property

Public Property Get Table() As TableManager.TableClass: Set Table = pTable: End Property

Public Property Get WorksheetName() As String: WorksheetName = pWorksheetName: End Property

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pControls = New TableManager.ControlsClass
    Set pEvents = New TableManager.EventsClass
End Sub                                          ' Class_Initialize

