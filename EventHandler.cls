VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormEvent As MSForms.UserForm
Attribute FormEvent.VB_VarHelpID = -1
Private pFormObj As Object

Public WithEvents SheetEvent As Worksheet
Attribute SheetEvent.VB_VarHelpID = -1
Private pSheet As Object

Public WithEvents ButtonEvent As MSForms.CommandButton
Attribute ButtonEvent.VB_VarHelpID = -1
Private pButtonObj As MSForms.CommandButton

Public WithEvents TextEvent As MSForms.TextBox
Attribute TextEvent.VB_VarHelpID = -1
Private pTextObj As MSForms.TextBox

Public WithEvents ComboEvent As MSForms.ComboBox
Attribute ComboEvent.VB_VarHelpID = -1
Private pComboObj As MSForms.ComboBox

Public WithEvents ImageEvent As MSForms.Image
Attribute ImageEvent.VB_VarHelpID = -1
Private pImageObj As Object

Private pTableObj As TableClass

Private pName As String

Private Sub ButtonEvent_Click()
    Dim Target As Range
    
    Select Case Me.ButtonEvent.Caption
    Case Is = "Copy To Table"
        pTableObj.PopulateTable pTableObj.ActiveTarget
    
    Case Is = "Cancel"
        pFormObj.Hide
        Exit Sub
    
    Case Is = "Next Row"
        pTableObj.NextRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(0, 0)
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Previous Row"
        pTableObj.PreviousRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(0, 0)
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Insert Above"
        pTableObj.ActiveTarget.ListObject.ListRows.Add pTableObj.ActiveRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(-1, 0)
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Insert Below"
        pTableObj.ActiveTarget.ListObject.ListRows.Add pTableObj.ActiveRow + 1
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(1, 0)
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Insert At Top"
        pTableObj.ActiveTarget.ListObject.ListRows.Add 1
        Set pTableObj.ActiveTarget = pTableObj.FirstCell
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Insert At End"
        pTableObj.ActiveTarget.ListObject.ListRows.Add
        Set pTableObj.ActiveTarget = pTableObj.LastCell
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Delete This Row"
        If pTableObj.LastRow = pTableObj.ActiveRow Then
            Set Target = pTableObj.ActiveTarget.Offset(-1, 0)
        Else
            Set Target = pTableObj.ActiveTarget.Offset(1, 0)
        End If
        pTableObj.ActiveTarget.ListObject.ListRows(pTableObj.ActiveRow).Delete
        Set pTableObj.ActiveTarget = Target
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Go To Top"
        Set pTableObj.ActiveTarget = pTableObj.FirstCell
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Is = "Go To Bottom"
        Set pTableObj.ActiveTarget = pTableObj.LastCell
        AllTbls(ActiveCellTableName).Form.PopulateForm pTableObj
    
    Case Else:
        Stop
    End Select
    
    PaintButtons
    
End Sub ' ButtonEvent_Click

Private Sub PaintButtons()
    Dim CtlIter As Variant
    
    For Each CtlIter In pFormObj.Controls
        With CtlIter
            If Left(.Name, 7) = "Command" Then
                EnableButton CtlIter
                LowLightControl CtlIter
                If .Caption = "Previous Row" And pTableObj.ActiveRow = 1 Then
                    DisableButton CtlIter
                End If
                If .Caption = "Next Row" And pTableObj.ActiveRow = pTableObj.LastRow Then
                    DisableButton CtlIter
                End If
            End If
        End With ' CtlIter
    Next CtlIter
End Sub

Private Sub ButtonEvent_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)
    
    PaintButtons
    
    HighLightButton pButtonObj
    
End Sub

Public Sub ButtonObj(ByVal Btn As MSForms.CommandButton)
    
    Set ButtonEvent = Btn
    Set pButtonObj = Btn
    pName = Btn.Name
End Sub

Private Sub ComboEvent_MouseMove( _
    ByVal Combo As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)
    
    Dim Ctl As Object

    For Each Ctl In pFormObj.Controls
        If Left(Ctl.Name, 5) <> "Image" Then
            If Left(Ctl.Name, 3) = "Com" And Ctl.ForeColor = vbWhite Then
                LowLightButton Ctl
            End If
        End If
    Next Ctl
    
    PaintButtons
    
End Sub

Public Sub ComboObj(ByVal Cmb As MSForms.ComboBox)
    Set ComboEvent = Cmb
    Set pComboObj = Cmb
    pName = Cmb.Name
End Sub

Private Sub FormEvent_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)
    
    Dim Ctl As Object

    For Each Ctl In pFormObj.Controls
        If Ctl.Enabled Then
            If Left(Ctl.Name, 5) <> "Image" Then
                If Left(Ctl.Name, 3) = "Com" Then
                    LowLightControl Ctl
                End If
            End If
        End If
    Next Ctl
    
    PaintButtons
    
End Sub

Private Sub SheetEvent_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim TableName As String
    
    Cancel = True

    TableName = ActiveCellTableName
    
    If TableName <> "" Then
        AllTbls(TableName).Form.PopulateForm AllTbls(TableName)
        AllTbls(TableName).Form(TableName).Show
    End If
    
    PaintButtons
    
End Sub

Public Property Set TableObj(ByVal Tbl As TableClass)
    Set pTableObj = Tbl
End Property

Public Property Set FormObj(ByVal Vbl As Object)
    Set pFormObj = Vbl
    Set Me.FormEvent = Vbl
End Property

Public Property Set ImageObj(ByVal Vbl As Object)
    Set pImageObj = Vbl
    Set Me.ImageEvent = Vbl
End Property

Public Property Get Name() As String: Name = pName: End Property

Public Sub TextObj(ByVal Txt As MSForms.TextBox)
    Set TextEvent = Txt
    Set pTextObj = Txt
    pName = Txt.Name
End Sub

Public Property Get Sheet() As Worksheet: Set Sheet = pSheet: End Property

Public Property Set SheetObj(ByVal Vbl As Object)
    Set pSheet = Vbl
    Set Me.SheetEvent = Vbl
End Property

