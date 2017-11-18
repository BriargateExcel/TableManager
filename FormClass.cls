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

Private Const Module_Name = "FormClass."

Private pFormName As String
Private pWorksheetName As String
Private pFormObj As Object
Private pTable As TableManager.TableClass
Private pControls As TableManager.ControlsClass
Private pEvents As TableManager.EventsClass

Private pHeight As Single       ' Form height
Private pWidth As Single        ' Form width
Private pTop As Single          ' Position of top of form
Private pLeft As Single         ' Position of form left edge
Private pCaption As String      ' Form caption
Private pEnabled As Boolean     ' Whether form is enabled
Private pPrevBtn As String
Private pNextBtn As String

' TODO: Calculate the gaps and multipliers based on font and font size
Private Const StandardGap = 12
Private Const ThinGap = 3
Private Const TextHeight = 18
Private Const ComboHeight = 18
Private Const ButtonHeight = 24
Private Const ButtonWidth = 72
Private Const FontSize = 8.25
Private Const WidthMultiplier = 6
Private Const TextMultiplier = 4.6

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pControls = New TableManager.ControlsClass
    Set pEvents = New TableManager.EventsClass
End Sub ' Class_Initialize

Public Property Get Name() As String: Name = pFormName: End Property
Public Property Let Name(ByVal FN As String): pFormName = FN: End Property

Public Property Get Table() As TableManager.TableClass: Set Table = pTable: End Property

Public Property Get WorksheetName() As String: WorksheetName = pWorksheetName: End Property

Public Property Get Caption() As String: Caption = pCaption: End Property

Public Property Get Controls() As TableManager.ControlsClass: Set Controls = pControls: End Property

Public Property Get Enabled() As Boolean: Enabled = pEnabled: End Property
Public Property Let Enabled(ByVal En As Boolean): pEnabled = En: End Property

Public Property Get FormObj() As MSForms.UserForm: Set FormObj = pFormObj: End Property
Public Property Set FormObj(ByVal FO As MSForms.UserForm): Set pFormObj = FO: End Property

Public Property Get Height() As Single: Height = pHeight: End Property

Public Property Get Left() As Single: Left = pLeft: End Property

Public Property Get Top() As Single: Top = pTop: End Property

Public Property Get Width() As Single: Width = pWidth: End Property

Private Sub BuildButton( _
    ByVal Cption As String, _
    ByVal Top As Single, _
    ByRef Lft As Single, _
    ByVal Tip As String, _
    ByVal Tbl As TableManager.TableClass)
    
    Dim Ctl As Control
    Dim Evt As TableManager.EventClass
    
    Const RoutineName = Module_Name & "BuildButton"
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
    
    If Ctl.Caption = "Previous Row" Then
        pPrevBtn = Ctl.Name
    End If
    If Ctl.Caption = "Next Row" Then
        pNextBtn = Ctl.Name
    End If

    pControls.Add Ctl
    
    Set Evt = New TableManager.EventClass
    Set Evt.ButtonObj = Ctl
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = Cption
    pEvents.Add Evt
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildButton

Private Sub BuildLabel( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxLabelWidth As Single, _
    ByRef Lbl As Control)

    Const RoutineName = Module_Name & "BuildLabel"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildLabel

Private Sub BuildDataLabel( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxFieldWidth As Single, _
    ByRef Lbl As Control)

    Const RoutineName = Module_Name & "BuildDataLabel"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildDataLabel

Private Sub BuildText( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxFieldWidth As Single, _
    ByRef Lbl As Control, _
    ByRef Tbl As TableManager.TableClass)

    Const RoutineName = Module_Name & "BuildText"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildText

Private Function BuildCaption( _
    ByVal InTitle As String, _
    ByVal InMessage As String _
    ) As String
    
    Dim Msg As String
    
    Msg = InTitle
    
    If InTitle <> "" Then Msg = Msg & ": "
    
    BuildCaption = Msg & InMessage
    
End Function

Private Sub BuildCombo( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxFieldWidth As Single, _
    ByRef Lbl As Control, _
    ByRef Tbl As TableManager.TableClass)

    Const RoutineName = Module_Name & "BuildCombo"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildCombo

Private Sub BuildDate( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxFieldWidth As Single, _
    ByRef Lbl As Control, _
    ByRef Tbl As TableManager.TableClass)

    Const RoutineName = Module_Name & "BuildDate"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildDate

Private Sub BuildWholeNumber( _
    ByRef Top As Single, _
    ByRef Field As Variant, _
    ByRef MaxFieldWidth As Single, _
    ByRef Lbl As Control, _
    ByRef Tbl As TableManager.TableClass)

    Const RoutineName = Module_Name & "BuildWholeNumber"
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
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildWholeNumber

Public Function BuildForm( _
    ByVal Tbl As TableManager.TableClass _
    ) As MSForms.UserForm
' TODO: refactor
    Dim TempForm As VBComponent
    
    Dim Field As Variant
    Dim Top As Single
    Dim Ctl As Control
    Dim Lbl As Control
    Dim MaxLabelWidth As Single
    Dim MaxFieldWidth As Single
    Dim MaxButtonWidth As Single
    Dim DataLeft As Single
    Dim LogoHeight As Single
    Dim LogoWidth As Single
    Dim Logo As Control
    Dim Evt As TableManager.EventClass
    Dim I As Long
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "BuildForm"
    On Error GoTo ErrorHandler
    
    Set pTable = Tbl
    pWorksheetName = Tbl.WorksheetName
    
    ' Create the UserForm
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    pFormName = TempForm.Name
    
    Set Tbl.Form = New TableManager.FormClass
    Tbl.Form.Name = pFormName
    Set Tbl.Form.FormObj = FormObj
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    
    Set Evt = New TableManager.EventClass
    Set Evt.TableObj = Tbl
    Set Evt.FormObj = pFormObj
    Evt.Name = pFormName
    pEvents.Add Evt
    
    With pFormObj
        
        ' Add a background texture
        If Dir(MainWorkbook.Path & "\texture.jpg") <> "" Then
            Set .Picture = LoadPicture(MainWorkbook.Path & "\texture.jpg")
        End If
        
        ' Add a Logo
        If Dir(MainWorkbook.Path & "\logo.jpg") <> "" Then
            Set Logo = .Controls.Add("Forms.Image.1")
            With Logo
                Set .Picture = LoadPicture(MainWorkbook.Path & "\logo.jpg")
                .PictureAlignment = fmPictureAlignmentTopRight
                .PictureSizeMode = fmPictureSizeModeZoom
                .BorderStyle = fmBorderStyleNone
                .BackStyle = fmBackStyleTransparent
                .AutoSize = True
                LogoHeight = .Height
                LogoWidth = .Width
            End With
        Else
            LogoHeight = 0
        End If
        Top = LogoHeight + 2 * StandardGap
        
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
                    
                Case xlValidateCustom: Stop

                Case xlValidateDate:
                    ' Set this up as a textbox
                    BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

                Case xlValidateDecimal
                    ' Set this up as a textbox
                    BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

                Case xlValidateTextLength
                    ' Set this up as a text field
                    BuildText Top, Field, MaxFieldWidth, Lbl, Tbl

                Case xlValidateTime: Stop

                Case xlValidateWholeNumber:
                    ' Set this up as a textbox
                    BuildWholeNumber Top, Field, MaxFieldWidth, Lbl, Tbl
                
                Case Else: Stop
                End Select
            End If ' Field is locked
            
        Next I
            
        ' Position the controls now that we know all the widths
        DataLeft = MaxLabelWidth + 2 * StandardGap
        For Each Ctl In pFormObj.Controls
            If Left(Ctl.Name, 3) = "lbl" Then
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
            
        ' Add First Row of Buttons
        DataLeft = StandardGap
        BuildButton "Validate Form", Top, DataLeft, _
            "Validate the contents of the Form", Tbl
        
        BuildButton "Copy To Table", Top, DataLeft, _
            "Copy the contents of the Form to the Table row", Tbl
        
        BuildButton "Next Row", Top, DataLeft, _
            "Move to next row of the Table", Tbl
        
        BuildButton "Previous Row", Top, DataLeft, _
            "Move to previous row of the Table", Tbl
        
        ' Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButton "Insert Above", Top, DataLeft, _
            "Insert a blank line above this Table row and move to it", Tbl
        
        BuildButton "Insert Below", Top, DataLeft, _
            "Insert a blank line below this Table row and move to it", Tbl
        
        BuildButton "Insert At Top", Top, DataLeft, _
            "Insert a blank line above the first row of the Table and move to it", Tbl
        
        MaxButtonWidth = DataLeft + ButtonWidth + 2 * StandardGap
        BuildButton "Insert At Bottom", Top, DataLeft, _
            "Insert a blank line below the last row of the Table and move to it", Tbl

        ' Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButton "Delete This Row", Top, DataLeft, _
            "Delete this row of the Table", Tbl
        
        BuildButton "Go To Top", Top, DataLeft, _
            "Move to the first row of the Table", Tbl
        
        BuildButton "Go To Bottom", Top, DataLeft, _
            "Move to the last row of the Table", Tbl

        BuildButton "Cancel", Top, DataLeft, _
            "Exit the Form without changing the Table any further", Tbl
        
        Top = Top + ButtonHeight + StandardGap
        
        ' Set Properties for TempForm now that we know all the heights and widths
        .Caption = Tbl.Name
        .Width = Application.Max(MaxLabelWidth + MaxFieldWidth + _
            4 * StandardGap, MaxButtonWidth)
        .Height = Top + 3 * StandardGap
        
        If Dir(MainWorkbook.Path & "\logo.jpg") <> "" Then
            Logo.Left = .Width - LogoWidth
            Logo.Left = 0
            Logo.Left = .InsideWidth - Logo.Width - StandardGap
            Logo.Top = StandardGap
        End If
        Set BuildForm = pFormObj
        
    End With ' pFormObj
    
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' BuildForm

Public Function NextButton() As String: NextButton = pNextBtn: End Function

Public Function PrevButton() As String: PrevButton = pPrevBtn: End Function

