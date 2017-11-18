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
Private pFormObj As Object 'MSForms.UserForm; when declared as UserFrom it has no height or width
                            ' When declared as an object, it has both height and width
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
Private Const TextMultiplier = 5

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
          Dim LogoImage As Control
          Dim Evt As TableManager.EventClass
          Dim I As Long
          
10        Debug.Assert Initializing

          Const RoutineName = Module_Name & "BuildForm"
20        On Error GoTo ErrorHandler
          
30        Set pTable = Tbl
40        pWorksheetName = Tbl.WorksheetName
          
          ' Create the UserForm
50        Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
60        pFormName = TempForm.Name
          
70        Set Tbl.Form = New TableManager.FormClass
80        Tbl.Form.Name = pFormName
90        Set pFormObj = VBA.UserForms.Add(TempForm.Name)
100       Set Tbl.Form.FormObj = pFormObj
          
110       Set Evt = New TableManager.EventClass
120       Set Evt.TableObj = Tbl
130       Set Evt.FormObj = pFormObj
140       Evt.Name = pFormName
150       pEvents.Add Evt
          
160       TableManager.Texture Tbl
          
170       Set LogoImage = TableManager.Logo(Tbl, LogoHeight, LogoWidth)
          
180       With pFormObj
              
190           Top = LogoHeight + 2 * StandardGap
              
              ' Build the field labels and fields
200           For I = 0 To Tbl.CellCount - 1
210               Set Field = Tbl.TableCells.Item(I)
                  
                  ' Build the field label
220               BuildLabel Top, Field, MaxLabelWidth, Lbl
                  
                  ' Build the data field
230               If Field.Locked Then
                      ' Set this up as a constant (label)
240                   BuildDataLabel Top, Field, MaxFieldWidth, Lbl
                      
250               Else
                      ' Field is unlocked
260                   Select Case Field.CellType
                      Case xlValidateInputOnly
                          ' Set this up as a text field
270                       BuildText Top, Field, MaxFieldWidth, Lbl, Tbl
                          
280                   Case xlValidateList
                          ' Set this up as a combobox
290                       BuildCombo Top, Field, MaxFieldWidth, Lbl, Tbl
                          
300                   Case xlValidateCustom: Stop

310                   Case xlValidateDate:
                          ' Set this up as a textbox
320                       BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

330                   Case xlValidateDecimal
                          ' Set this up as a textbox
340                       BuildDate Top, Field, MaxFieldWidth, Lbl, Tbl

350                   Case xlValidateTextLength
                          ' Set this up as a text field
360                       BuildText Top, Field, MaxFieldWidth, Lbl, Tbl

370                   Case xlValidateTime: Stop

380                   Case xlValidateWholeNumber:
                          ' Set this up as a textbox
390                       BuildWholeNumber Top, Field, MaxFieldWidth, Lbl, Tbl
                      
400                   Case Else: Stop
410                   End Select
420               End If ' Field is locked
                  
430           Next I
                  
              ' Position the controls now that we know all the widths
440           DataLeft = MaxLabelWidth + 2 * StandardGap
450           For Each Ctl In pFormObj.Controls
460               If Left(Ctl.Name, 3) = "lbl" Then
470                   Ctl.Width = MaxLabelWidth
480                   Ctl.Left = StandardGap
490               Else
500                   If Ctl.Name <> "Filler" Then
510                       Ctl.Width = MaxFieldWidth
520                       Ctl.Left = DataLeft
530                   End If
540               End If
550           Next Ctl
              
560           Top = Top + StandardGap
              
              ' TODO: Arrange the buttons based on the width of the form
              ' A wide form would have fewer rows of buttons than a narrow form
                  
              ' Add First Row of Buttons
570           DataLeft = StandardGap
580           BuildButton "Validate Form", Top, DataLeft, _
                  "Validate the contents of the Form", Tbl
              
590           BuildButton "Copy To Table", Top, DataLeft, _
                  "Copy the contents of the Form to the Table row", Tbl
              
600           BuildButton "Next Row", Top, DataLeft, _
                  "Move to next row of the Table", Tbl
              
610           BuildButton "Previous Row", Top, DataLeft, _
                  "Move to previous row of the Table", Tbl
              
              ' Second row of buttons
620           Top = Top + ButtonHeight + StandardGap
630           DataLeft = StandardGap
640           BuildButton "Insert Above", Top, DataLeft, _
                  "Insert a blank line above this Table row and move to it", Tbl
              
650           BuildButton "Insert Below", Top, DataLeft, _
                  "Insert a blank line below this Table row and move to it", Tbl
              
660           BuildButton "Insert At Top", Top, DataLeft, _
                  "Insert a blank line above the first row of the Table and move to it", Tbl
              
670           BuildButton "Insert At Bottom", Top, DataLeft, _
                  "Insert a blank line below the last row of the Table and move to it", Tbl

              ' Third row of buttons
680           MaxButtonWidth = DataLeft + ButtonWidth + 2 * StandardGap
690           Top = Top + ButtonHeight + StandardGap
700           DataLeft = StandardGap
710           BuildButton "Delete This Row", Top, DataLeft, _
                  "Delete this row of the Table", Tbl
              
720           BuildButton "Go To Top", Top, DataLeft, _
                  "Move to the first row of the Table", Tbl
              
730           BuildButton "Go To Bottom", Top, DataLeft, _
                  "Move to the last row of the Table", Tbl

740           BuildButton "Cancel", Top, DataLeft, _
                  "Exit the Form without changing the Table any further", Tbl
              
750           Top = Top + ButtonHeight + StandardGap
              
              ' Set Properties for TempForm now that we know all the heights and widths
760           .Caption = Tbl.Name
770           .Width = Application.Max( _
                  MaxLabelWidth + MaxFieldWidth + 4 * StandardGap, _
                  MaxButtonWidth)
780           .Height = Top + 3 * StandardGap
              
790           If Not LogoImage Is Nothing Then
800               LogoImage.Left = .Width - LogoWidth
810               LogoImage.Left = 0
820               LogoImage.Left = .InsideWidth - LogoImage.Width - StandardGap
830               LogoImage.Top = StandardGap
'840               Set pFormObj.Picture = LogoImage
850           End If
              
860       End With ' pFormObj
          
870       Set BuildForm = pFormObj
              
Done:
880       Exit Function
ErrorHandler:
          
890       RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' BuildForm

Public Function NextButton() As String: NextButton = pNextBtn: End Function

Public Function PrevButton() As String: PrevButton = pPrevBtn: End Function

