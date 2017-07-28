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

Private pHeight As Single       ' Form height
Private pWidth As Single        ' Form width
Private pTop As Single          ' Position of top of form
Private pLeft As Single         ' Position of form left edge
Private pCaption As String      ' Form caption
Private pEnabled As Boolean     ' Whether form is enabled
Private pName As String
Private pFormObj As Object
Private pTable As TableClass

Private Frm As EventHandler

Private Ctls As ControlsClass

Private Const StandardGap = 12
Private Const ThinGap = 3
Private Const TextHeight = 15
Private Const ComboHeight = 15
Private Const ButtonHeight = 24
Private Const ButtonWidth = 72
    


' Public methods

Public Function BuildForm( _
    ByVal Tbl As TableClass _
    ) As MSForms.UserForm

'   Description: Creates a new form

'   Inputs:
'   Tbl         The table corresponding to this form
'   Outputs:
'   Me          A form
'   Requisites: modGeneral
'   Notes:      ' Any notes
'   Example:    ' How to call this routine
'   History
'   05/08/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "BuildForm"
    Dim TempForm As VBComponent
    
    Dim Field As CellClass
    Dim Top As Single
    Dim Ctl As Control
    Dim Lbl As Control
    Dim MaxLabelWidth As Single
    Dim MaxFieldWidth As Single
    Dim DataLeft As Single
    Dim ValListItem As Variant
    Dim LogoHeight As Single
    Dim LogoWidth As Single
    Dim Logo As Control
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    pName = Tbl.Name
'    Set Tbl.Target = Target
    Set pTable = Tbl
    
'   Create the UserForm
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    Set Frm = New EventHandler
    Set Frm.FormObj = pFormObj
    Set Frm.FormEvent = pFormObj
    
    With pFormObj
        
    '   Add a background texture
        If Dir(Application.ActiveWorkbook.Path & "\texture.jpg") <> "" Then
            Set .Picture = LoadPicture(Application.ActiveWorkbook.Path & "\texture.jpg")
        End If
        
    '   Add a Logo
        If Dir(Application.ActiveWorkbook.Path & "\logo.jpg") <> "" Then
            Set Logo = .Controls.Add("Forms.Image.1")
            With Logo
                Set .Picture = LoadPicture(Application.ActiveWorkbook.Path & "\logo.jpg")
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
        
        For Each Field In Tbl.TableCells
'           Build the field labels
            Set Lbl = .Controls.Add("Forms.Label.1", "lbl" & Field.Name, True)
            Lbl.Top = Top
            Lbl.Left = StandardGap
            Lbl.Caption = Field.HeaderText
            Lbl.TextAlign = fmTextAlignRight
            Lbl.WordWrap = False
            Lbl.BackColor = LabelBackGround
            Lbl.ForeColor = LabelFont
            If Lbl.Width > MaxLabelWidth Then MaxLabelWidth = Lbl.Width
            
'           Build the data fields
            If Field.Locked Then
'           Set this up as a constant (label)
                Set Ctl = .Controls.Add("Forms.Label.1", "fld" & Field.Name, True)
                With Ctl
                    .Top = Top
                    .Height = TextHeight + ThinGap
                    Lbl.Height = TextHeight + ThinGap
                    Top = Top + TextHeight + ThinGap
                    .WordWrap = False
                    .BackColor = LabelBackGround
                    .ForeColor = LabelFont
                    .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                    If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                End With
                Set Field.FormControl = Ctl
            Else
'               Field is unlocked
                Select Case Field.CellType
                Case xlValidateInputOnly
'                   Set this up as a text field
                    Set Ctl = .Controls.Add("Forms.TextBox.1", "fld" & Field.Name, True)
                    
                    With Ctl
                        .Top = Top
                        .Width = Field.ColumnWidth * 4
                        If Field.Wrap Then
                        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                            .Height = 3 * TextHeight
                            Lbl.Height = 3 * TextHeight
                            Top = Top + 3 * TextHeight
                            .WordWrap = True
                            .MultiLine = True
                            .ScrollBars = fmScrollBarsVertical
                            .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                        Else
                            .Height = TextHeight
                            Lbl.Height = TextHeight
                            Top = Top + TextHeight
                            .WordWrap = False
                            .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                        End If
                    End With
                    
                    Frm.TextObj Ctl
                    
                Case xlValidateList
'                   Set this up as a combobox
                    Set Ctl = .Controls.Add("Forms.ComboBox.1", "cmb" & Field.Name, True)
                    With Ctl
                        .Top = Top
                        .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                        .Height = ComboHeight
                        Lbl.Height = ComboHeight
                        Top = Top + ComboHeight
                        .Width = Field.ColumnWidth * 4
                        If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                        For Each ValListItem In Field.ValidationList
                            .AddItem CStr(ValListItem)
                        Next ValListItem
                        Top = Top
                    End With
                    Frm.ComboObj Ctl
                
'                Case xlValidateCustom: Stop
'
'                Case xlValidateDate: Stop
'
'                Case xlValidateDecimal: Stop
'
'                Case xlValidateTextLength: Stop
'
'                Case xlValidateTime: Stop
'
'                Case xlValidateWholeNumber: Stop
'
'                Case Else: Stop
                End Select
                
                Set Field.FormControl = Ctl
            End If ' Field is locked
            
        Next Field
            
'        PopulateForm Tbl
        
        Set Ctl = .Controls.Add("Forms.Label.1", "Filler", True)
        With Ctl
            .Left = MaxLabelWidth + StandardGap
            .Top = LogoHeight + 2 * StandardGap
            .Width = StandardGap
            .Height = Top - .Top
            .BackColor = LabelBackGround
            .ForeColor = LabelFont
        End With ' Ctl

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
            
'       Add First Row of Buttons
        DataLeft = StandardGap
        BuildButton "Copy To Table", Top, DataLeft, "Copy the contents of the form to the table row", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Cancel", Top, DataLeft, "Exit the form without changing the table", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Next Row", Top, DataLeft, "Move to next row of the table", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Previous Row", Top, DataLeft, "Move to previous row of the table", Tbl
        
'       Second row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButton "Insert Above", Top, DataLeft, "Insert a blank line above this table row and move to it", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Insert Below", Top, DataLeft, "Insert a blank line below this table row and move to it", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Insert At Top", Top, DataLeft, "Insert a blank line above the first row of the table and move to it", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Insert At End", Top, DataLeft, "Insert a blank line below the last row of the table and move to it", Tbl

'       Third row of buttons
        Top = Top + ButtonHeight + StandardGap
        DataLeft = StandardGap
        BuildButton "Delete This Row", Top, DataLeft, "Delete this row of the table", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Go To Top", Top, DataLeft, "Move to the first row of the table", Tbl
        
        DataLeft = DataLeft + ButtonWidth + StandardGap
        BuildButton "Go To Bottom", Top, DataLeft, "Move to the last row of the table", Tbl

        Top = Top + ButtonHeight + StandardGap
        
'       Set Properties for TempForm
        .Caption = Tbl.Name
        .Width = MaxLabelWidth + MaxFieldWidth + 4 * StandardGap
        .Height = Top + 3 * StandardGap
        If Dir(Application.ActiveWorkbook.Path & "\logo.jpg") <> "" Then
            Logo.Left = .Width - LogoWidth
            Logo.Left = 0
            Logo.Left = .InsideWidth - Logo.Width - StandardGap
            Logo.Top = StandardGap
        End If
        Set BuildForm = pFormObj
        
    End With ' pFormObj
    
    Set Tbl.Form = Me
    
ErrHandler:
    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Is = 2015
            On Error GoTo ErrHandler
            Resume Next
        Case Else:
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Function ' BuildForm

Private Sub BuildButton( _
    ByVal Cption As String, _
    ByVal Top As Single, _
    ByVal Lft As Single, _
    ByVal Tip As String, _
    ByVal Tbl As TableClass)
    
    Dim Ctl As Control
    
    Set Ctl = pFormObj.Controls.Add("Forms.CommandButton.1")
    With Ctl
        .Caption = Cption
        .Top = Top
        .Left = Lft
        .Height = ButtonHeight
        .Width = ButtonWidth
        .BackColor = ButtonNothingBackGround
        .ForeColor = ButtonNothingFont
        .ControlTipText = Tip
    End With
    Ctls.Add Ctl
    Ctls.Item(Ctl.Name).ButtonObj Ctl
    Set Ctls.Item(Ctl.Name).FormObj = pFormObj
    Set Ctls.Item(Ctl.Name).TableObj = Tbl
End Sub


Public Function PopulateForm( _
    ByVal Tbl As TableClass _
    ) As MSForms.UserForm

'   Description: Populates the controls on a form with the data in the row
'   Inputs:
'   Tbl         The table corresponding to this form
'   Outputs:
'   Me          A form
'   Requisites: modGeneral
'   Notes:      ' Any notes
'   Example:    ' How to call this routine
'   History
'   05/08/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "PopulateForm"
    
    Dim Field As CellClass
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim THRow As Long: THRow = Tbl.ActiveTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = Tbl.ActiveTarget.Row
    Dim DBRow As Long: DBRow = TRow - THRow
    Dim DBCol As Long
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    
'   Populate the data fields
    For Each Field In Tbl.TableCells
        DBCol = Application.Match(Field.HeaderText, Tbl.ActiveTarget.ListObject.HeaderRowRange, 0)
        If Field.Locked Then
            Field.ControlValue = DBRange(DBRow, DBCol)
        Else
            Field.ControlValue = DBRange(DBRow, DBCol)
        End If
        Field.FormControl = DBRange(DBRow, DBCol)
    Next Field
 
ErrHandler:
    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Else:
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Function ' PopulateForm

Public Sub Show()
    pFormObj.Show
End Sub

Public Sub CloseForm()
    Unload pFormObj
End Sub

' Properties

Public Property Get Height() As Single: Height = pHeight: End Property

Public Property Get Width() As Single: Width = pWidth: End Property

Public Property Get Top() As Single: Top = pTop: End Property

Public Property Get Left() As Single: Left = pLeft: End Property

Public Property Get Caption() As String: Caption = pCaption: End Property

Public Property Get FormObj() As MSForms.UserForm: Set FormObj = pFormObj: End Property

Public Property Get Enabled() As Boolean:  Enabled = pEnabled: End Property
Public Property Let Enabled(ByVal En As Boolean):  pEnabled = En: End Property

Public Property Get Name() As String: Name = pName: End Property
Public Property Let Name(ByVal Vbl As String): pName = Vbl: End Property

Private Sub Class_Initialize()
    Set Ctls = New ControlsClass
End Sub

Public Property Get Table() As TableClass: Set Table = pTable: End Property
Public Property Set Table(ByVal Vbl As TableClass): Set pTable = Vbl: End Property

