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

Private frm As EventHandler

Private Ctls As ControlsClass


' Public methods

Public Function BuildForm( _
    ByVal Tbl As TableClass, _
    Target As Range) As MSForms.UserForm

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
    Set Tbl.Target = Target
    Set pTable = Tbl
    
'   Create the UserForm
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    Set frm = New EventHandler
    Set frm.FormObj = pFormObj
    Set frm.FormEvent = pFormObj
    
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
            LogoHeight = .Height + 12
            LogoWidth = .Width
        End With
        Ctls.Add Logo
        Top = LogoHeight
    Else
        Top = 6
    End If
    
'       Build the field labels
        For Each Field In Tbl.TableCells
            Set Ctl = .Controls.Add("Forms.Label.1", "lbl" & Field.Name, True)
            With Ctl
                .Top = Top
                .Left = 6
                .Height = 18
                If Field.Wrap Then
                    .Height = 54 ' data field is three lines high
                    Top = Top + 54 ' 48 for data field height + 6 for gap between labels
                Else
                    .Height = 18 ' data field is one line high
                    Top = Top + 18 ' 12 for data field height + 6 for gap between labels
                End If
                .Caption = Field.HeaderText
                .TextAlign = fmTextAlignRight
                .WordWrap = False
                .BackColor = LabelBackGround
                .ForeColor = LabelFont
                If .Width > MaxLabelWidth Then MaxLabelWidth = .Width
            End With
            Ctls.Add Ctl
        Next Field
        
        DataLeft = MaxLabelWidth + 12
        If Dir(Application.ActiveWorkbook.Path & "\logo.jpg") <> "" Then
            Top = LogoHeight
        Else
            Top = 6
        End If
        
'       Build the data fields
        For Each Field In Tbl.TableCells
            If Field.Locked Then
'               Set this up as a constant (label)
                Set Ctl = .Controls.Add("Forms.Label.1", "fld" & Field.Name, True)
                With Ctl
                    .Top = Top
                    .Left = DataLeft
                    .Height = 18 ' data field is one line high
                    .WordWrap = False
                    .BackColor = LabelBackGround
                    .ForeColor = LabelFont
                    .ControlTipText = Field.InputTitle & vbCr & Field.InputMessage
                    If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                End With
                Set Field.FormControl = Ctl
                Ctls.Add Ctl
            Else
                ' Field is unlocked
                Select Case Field.CellType
                    Case xlValidateCustom: Stop
                    
                    Case xlValidateDate: Stop
                    
                    Case xlValidateDecimal: Stop
                    
                    Case xlValidateInputOnly
'                       Set this up as a text field
                        Set Ctl = .Controls.Add("Forms.TextBox.1", "fld" & Field.Name, True)
                        
                        With Ctl
                            .Top = Top
                            .Left = DataLeft
                            .Width = Field.ColumnWidth * 4 + 20
                            If Field.Wrap Then
                            If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                                .Height = 54
                                .WordWrap = True
                                .MultiLine = True
                                .ScrollBars = fmScrollBarsVertical
                                .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                            Else
                                .WordWrap = False
                                .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                            End If
                        End With
                        
                        Ctls.Add Ctl
                        Ctls.Item(Ctl.Name).TextObj Ctl
                        frm.TextObj Ctl
                        
                    Case xlValidateList
'                       Set this up as a combobox
                        Set Ctl = .Controls.Add("Forms.ComboBox.1", "cmb" & Field.Name, True)
                        With Ctl
                            .Top = Top
                            .Left = DataLeft
                            .ControlTipText = Field.InputTitle & ": " & Field.InputMessage
                            .Width = Field.ColumnWidth * 4
                            If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                            For Each ValListItem In Field.ValidationList
                                .AddItem CStr(ValListItem)
                            Next ValListItem
                        End With
                        Ctls.Add Ctl
                        Ctls.Item(Ctl.Name).ComboObj Ctl
                        frm.ComboObj Ctl
                    
                    Case xlValidateTextLength: Stop
                    
                    Case xlValidateTime: Stop
                    
                    Case xlValidateWholeNumber: Stop
                    
                    Case Else
                        Stop
                End Select
                Ctl.BackColor = vbWhite
                Ctl.ForeColor = vbBlack
                Set Field.FormControl = Ctl
                Set Ctls.Item(Ctl.Name).FormObj = pFormObj
            End If ' Field is locked
            
            If Field.Wrap Then
                Top = Top + 54 ' 48 for data field height + 6 for gap between labels
            Else
                Top = Top + 18 ' 12 for data field height + 6 for gap between labels
            End If
        Next Field
        
        PopulateForm Tbl, Target
 
        For Each Ctl In pFormObj.Controls
            If Left(Ctl.Name, 3) = "lbl" Then
                Ctl.Width = MaxLabelWidth
            Else
                Ctl.Width = MaxFieldWidth
            End If
        Next Ctl
        
        Top = Top + 6
            
'       Add First Row of Buttons
        BuildButton "Copy To Table", Top, 20, "Copy the contents of the form to the table row", Tbl
        BuildButton "Cancel", Top, 98, "Exit the form without changing the table", Tbl
        BuildButton "Next Row", Top, 176, "Move to next row of the table", Tbl
        BuildButton "Previous Row", Top, 254, "Move to previous row of the table", Tbl
        
'       Second row of buttons
        Top = Top + 30
        BuildButton "Insert Above", Top, 20, "Insert a blank line above this table row and move to it", Tbl
        BuildButton "Insert Below", Top, 98, "Insert a blank line below this table row and move to it", Tbl
        BuildButton "Insert At Top", Top, 176, "Insert a blank line above the first row of the table and move to it", Tbl
        BuildButton "Insert At End", Top, 254, "Insert a blank line below the last row of the table and move to it", Tbl

'       Third row of buttons
        Top = Top + 30
        BuildButton "Delete This Row", Top, 20, "Delete this row of the table", Tbl
        BuildButton "Go To Top", Top, 98, "Move to the first row of the table", Tbl
        BuildButton "Go To Bottom", Top, 176, "Move to the last row of the table", Tbl

'       Set Properties for TempForm
        .Caption = Tbl.Name
        .Width = MaxLabelWidth + MaxFieldWidth + 30
        .Height = Top + 60
        If Dir(Application.ActiveWorkbook.Path & "\logo.jpg") <> "" Then
            Logo.Left = .Width - LogoWidth
            Logo.Left = 0
            Logo.Left = .InsideWidth - Logo.Width - 6
            Logo.Top = 6
        End If
        Set BuildForm = pFormObj
        
    End With ' pFormObj
    
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
        .Height = 24
        .Width = 72
        .BackColor = ButtonNothingBackGround
        .ForeColor = ButtonNothingFont
        .ControlTipText = Tip
    End With
    Ctls.Add Ctl
    Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
    Set Ctls.Item(Ctl.Name).FormObj = pFormObj
    Set Ctls.Item(Ctl.Name).TableObj = Tbl
End Sub


Public Function PopulateForm( _
    ByVal Tbl As TableClass, _
    Target As Range) As MSForms.UserForm

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
    Dim DBRange As Range: Set DBRange = Target.ListObject.DataBodyRange
    Dim THRow As Long: THRow = Target.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = Target.Row
    Dim DBrow As Long: DBrow = TRow - THRow
    Dim DBCol As Long
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    
'   Populate the data fields
    For Each Field In Tbl.TableCells
        DBCol = Application.Match(Field.HeaderText, Target.ListObject.HeaderRowRange, 0)
        If Field.Locked Then
            Field.ControlValue = DBRange(DBrow, DBCol)
        Else
            Field.ControlValue = DBRange(DBrow, DBCol)
        End If
        Field.FormControl = DBRange(DBrow, DBCol)
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

Public Property Get ControlsList() As Variant: ControlsList = Ctls.ListCtls: End Property

Private Sub Class_Initialize()
    Set Ctls = New ControlsClass
End Sub

Public Property Get Table() As TableClass: Set Table = pTable: End Property
Public Property Set Table(ByVal Vbl As TableClass): Set pTable = Vbl: End Property

