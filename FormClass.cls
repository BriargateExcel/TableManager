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
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    pName = Tbl.Name
    Set Tbl.Target = Target
    Set pTable = Tbl
    
'   Create the UserForm
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    Set pFormObj = VBA.UserForms.Add(TempForm.Name)
    Set Frm = New EventHandler
    Set Frm.FormObj = pFormObj
    Set Frm.FormEvent = pFormObj
    
    With pFormObj
        
'       Build the field labels
        Top = 6
        For Each Field In Tbl.TableCells
            Set Ctl = .Controls.Add("Forms.Label.1", "lbl" & Field.Name, True)
            With Ctl
                .Top = Top
                .Left = 6
                .Height = 12
                If Field.Wrap Then
                    .Height = 36 ' data field is three lines high
                    Top = Top + 54 ' 48 for data field height + 6 for gap between labels
                Else
                    .Height = 12 ' data field is one line high
                    Top = Top + 18 ' 12 for data field height + 6 for gap between labels
                End If
                .Caption = Field.HeaderText
                .TextAlign = fmTextAlignRight
                .WordWrap = False
                .AutoSize = True
                If .Width > MaxLabelWidth Then MaxLabelWidth = .Width

            End With
            Ctls.Add Ctl
        Next Field
        
        DataLeft = MaxLabelWidth + 12
        Top = 6
        
'       Build the data fields
        For Each Field In Tbl.TableCells
            If Field.Locked Then
'               Set this up as a constant (label)
                Set Ctl = .Controls.Add("Forms.Label.1", "fld" & Field.Name, True)
                With Ctl
                    .Top = Top
                    .Left = DataLeft
                    .Height = 12 ' data field is one line high
                    .WordWrap = False
                    .AutoSize = True
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
                            .Width = Application.Max(Field.ColumnWidth * 5, 60)
                            If Field.Wrap Then
                            If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                                .Height = 54
                                .WordWrap = True
                                .MultiLine = True
                            Else
                                .WordWrap = False
                            End If
                        End With
                        
                        Ctls.Add Ctl
                        Ctls.Item(Ctl.Name).TextObj Ctl, Ctl.Name
                        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
                        Frm.TextObj Ctl, Ctl.Name
                        
                    Case xlValidateList
'                       Set this up as a combobox
                        Set Ctl = .Controls.Add("Forms.ComboBox.1", "cmb" & Field.Name, True)
                        With Ctl
                            .Top = Top
                            .Left = DataLeft
                            .Width = Application.Max(Field.ColumnWidth * 5, 60)
                            If .Width > MaxFieldWidth Then MaxFieldWidth = .Width
                            For Each ValListItem In Field.ValidationList
                                .AddItem CStr(ValListItem)
                            Next ValListItem
                        End With
                        Ctls.Add Ctl
                        Ctls.Item(Ctl.Name).ComboObj Ctl, Ctl.Name
                        Frm.ComboObj Ctl, Ctl.Name
                    
                    Case xlValidateTextLength: Stop
                    
                    Case xlValidateTime: Stop
                    
                    Case xlValidateWholeNumber: Stop
                    
                    Case Else
                        Stop
                End Select
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
            
'       Add Copy To Table Button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Copy To Table"
            .Top = Top
            .Left = 20
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Cancel button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Cancel"
            .Top = Top
            .Left = 98
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Next Row button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Next Row"
            .Top = Top
            .Left = 176
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Previous Row button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Previous Row"
            .Top = Top
            .Left = 254
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl

'       Second row of buttons
        Top = Top + 30
        
'       Add Insert Above button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Insert Above"
            .Top = Top
            .Left = 20
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Insert Below button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Insert Below"
            .Top = Top
            .Left = 98
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Insert at Top button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Insert at Top"
            .Top = Top
            .Left = 176
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Insert At End button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Insert At End"
            .Top = Top
            .Left = 254
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Third row of buttons
        Top = Top + 30
        
'       Add Delete button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Delete This Row"
            .Top = Top
            .Left = 20
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Go To Top button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Go To Top"
            .Top = Top
            .Left = 98
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Add Go To Bottom button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Go To Bottom"
            .Top = Top
            .Left = 176
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Set Properties for TempForm
        .Caption = Tbl.Name
        .Width = Application.Max(MaxLabelWidth + MaxFieldWidth + 30, 666)
        .Height = Top + 90
        Set BuildForm = pFormObj
        
'       Add Delete button
        Set Ctl = .Controls.Add("Forms.CommandButton.1")
        With Ctl
            .Caption = "Delete This Row"
            .Top = Top
            .Left = 20
            .Height = 24
            .Width = 72
        End With
        Ctls.Add Ctl
        Ctls.Item(Ctl.Name).ButtonObj Ctl, Ctl.Name
        Set Ctls.Item(Ctl.Name).FormObj = pFormObj
        Set Ctls.Item(Ctl.Name).TableObj = Tbl
        
'       Set Properties for TempForm
        .Caption = Tbl.Name
        .Width = Application.Max(MaxLabelWidth + MaxFieldWidth + 30, 666)
        .Height = Top + 90
        Set BuildForm = pFormObj
        
    End With
    
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

Public Function PopulateForm(ByVal Tbl As TableClass, _
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
    Dim TempForm As VBComponent
    
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

