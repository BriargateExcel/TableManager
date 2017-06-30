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

Public WithEvents ButtonEvent As MSForms.CommandButton
Attribute ButtonEvent.VB_VarHelpID = -1
Private pButtonObj As MSForms.CommandButton

Public WithEvents TextEvent As MSForms.TextBox
Attribute TextEvent.VB_VarHelpID = -1
Private pTextObj As MSForms.TextBox

Public WithEvents ComboEvent As MSForms.ComboBox
Attribute ComboEvent.VB_VarHelpID = -1
Private pComboObj As MSForms.ComboBox

Private pTableObj As TableClass

Private pName As String
Private pTableName As String

Private Sub ButtonEvent_Click()
    Dim Target As Range

    If Me.ButtonEvent.Caption = "Copy To Table" Then
        pTableObj.PopulateTable pTableObj.ActiveTarget
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Cancel" Then
        Unload pFormObj
        DeleteForm
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Next Row" Then
        pTableObj.NextRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(0, 0)
        ActiveForm.PopulateForm pTableObj, pTableObj.ActiveTarget
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Previous Row" Then
        pTableObj.PreviousRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(0, 0)
        ActiveForm.PopulateForm pTableObj, pTableObj.ActiveTarget
        Exit Sub
    End If

    ' Position=3 inserts a row above row 3
    ' Position=2 inserts a row above row 2, etc
    ' Position=1 inserts a row at the beginning of the table
    ' .Add without specifying Position adds a row at the bottom of the table
    ' Position <=0 fails for subscript out of range

    If Me.ButtonEvent.Caption = "Insert Above" Then
        pTableObj.ActiveTarget.ListObject.ListRows.Add pTableObj.ActiveRow
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(-1, 0)
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Insert Below" Then
        pTableObj.ActiveTarget.ListObject.ListRows.Add pTableObj.ActiveRow + 1
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(1, 0)
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Insert at Top" Then
        pTableObj.ActiveTarget.ListObject.ListRows.Add 1
        Set pTableObj.ActiveTarget = pTableObj.ActiveTarget.Offset(1, 0)
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Insert At End" Then
        pTableObj.ActiveTarget.ListObject.ListRows.Add
        Exit Sub
    End If

    If Me.ButtonEvent.Caption = "Delete This Row" Then
        If pTableObj.LastRow = pTableObj.ActiveRow Then
            Set Target = pTableObj.ActiveTarget.Offset(-1, 0)
        Else
            Set Target = pTableObj.ActiveTarget.Offset(1, 0)
        End If
        pTableObj.ActiveTarget.ListObject.ListRows(pTableObj.ActiveRow).Delete
        Set pTableObj.ActiveTarget = Target
        ActiveForm.PopulateForm pTableObj, pTableObj.ActiveTarget
        Exit Sub
    End If
    
    If Me.ButtonEvent.Caption = "Go To Top" Then
        Set pTableObj.ActiveTarget = pTableObj.FirstCell
        ActiveForm.PopulateForm pTableObj, pTableObj.ActiveTarget
    End If

    If Me.ButtonEvent.Caption = "Go To Bottom" Then
        Set pTableObj.ActiveTarget = pTableObj.LastCell
        ActiveForm.PopulateForm pTableObj, pTableObj.ActiveTarget
    End If

End Sub

Private Sub FormEvent_MouseMove(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
Dim Ctl As Object

    For Each Ctl In pFormObj.Controls
        If Left(Ctl.Name, 3) = "fld" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "cmb" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "Com" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        End If
    Next Ctl
    
End Sub

Public Property Set FormObj(ByVal Vbl As Object)
    Set pFormObj = Vbl
    Set Me.FormEvent = Vbl
End Property

Public Property Set TableObj(ByVal Vbl As Object)
    Set pTableObj = Vbl
End Property

Private Sub ButtonEvent_MouseMove(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    With pButtonObj
        .ForeColor = vbWhite
        .BackColor = vbBlack
    End With
    
End Sub

Public Sub ButtonObj(ByVal Btn As MSForms.CommandButton, _
                                ByVal BtnName As String)
    Set ButtonEvent = Btn
    Set pButtonObj = Btn
    pName = BtnName
End Sub

Private Sub TextEvent_MouseMove(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
Dim Ctl As Object

    For Each Ctl In pFormObj.Controls
        If Left(Ctl.Name, 3) = "fld" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "cmb" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "Com" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .Forelor = vbBlack
                .BackColor = vbWhite
            End With
        End If
        If Ctl.Name = Me.Name Then
            With pTextObj
                .ForeColor = vbWhite
                
                .BackColor = vbBlack
            End With
        End If
    Next Ctl

End Sub


Public Sub TextObj(ByVal Txt As MSForms.TextBox, _
                            ByVal TxtName As String)
    Set TextEvent = Txt
    Set pTextObj = Txt
    pName = TxtName
End Sub

Private Sub ComboEvent_MouseMove(ByVal Combo As Integer, _
    ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
Dim Ctl As Object

    For Each Ctl In pFormObj.Controls
        If Left(Ctl.Name, 3) = "fld" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "cmb" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        ElseIf Left(Ctl.Name, 3) = "Com" And Ctl.ForeColor = vbWhite Then
            With Ctl
                .ForeColor = vbBlack
                .BackColor = vbWhite
            End With
        End If
        If Ctl.Name = Me.Name Then
            With pComboObj
                .ForeColor = vbWhite
                .BackColor = vbBlack
            End With
        End If
    Next Ctl
    
End Sub

Public Sub ComboObj(ByVal Cmb As MSForms.ComboBox, _
                                ByVal CmbName As String)
    Set ComboEvent = Cmb
    Set pComboObj = Cmb
    pName = CmbName
End Sub

Public Property Get Name() As String: Name = pName: End Property
