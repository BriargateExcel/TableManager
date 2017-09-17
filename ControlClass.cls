VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "ControlClass."

Public WithEvents TxtCtl As MSForms.TextBox
Attribute TxtCtl.VB_VarHelpID = -1
Private pTxtObj As MSForms.TextBox

Public WithEvents BtnCtl As MSForms.CommandButton
Attribute BtnCtl.VB_VarHelpID = -1
Private pBtnObj As MSForms.CommandButton

Public WithEvents CmbCtl As MSForms.ComboBox
Attribute CmbCtl.VB_VarHelpID = -1
Private pCmbObj As MSForms.ComboBox

Private pEventHandler As TableManager.EventHandler

Private pName As String
Private pField As TableManager.CellClass

Public Sub ComboInit( _
    ByVal Tbl As TableManager.TableClass, _
    ByVal Frm As MSForms.UserForm, _
    ByVal Field As TableManager.CellClass, _
    ByVal Cmb As MSForms.ComboBox)
    
    Me.Name = Cmb.Name
    Set Me.CmbCtl = Cmb
    Set pCmbObj = Cmb
    Set Field.Ctl = Me
    Set pField = Field
    
    Set pEventHandler = New TableManager.EventHandler
    pEventHandler.ComboObj Cmb
    Set pEventHandler.TableObj = Tbl
    Set pEventHandler.FormObj = Frm
    Set pEventHandler.ComboEvent = Cmb
End Sub

Public Sub DisableControl()
    pEventHandler.DisableEvents
End Sub

Public Sub EnableControl()
    pEventHandler.EnableEvents
End Sub

Public Sub TextInit( _
    ByVal Tbl As TableManager.TableClass, _
    ByVal Frm As MSForms.UserForm, _
    ByVal Field As TableManager.CellClass, _
    ByVal Txt As MSForms.TextBox)
    
    Me.Name = Txt.Name
    Set Me.TxtCtl = Txt
    Set pTxtObj = Txt
    Set Field.Ctl = Me
    Set pField = Field
    
    Set pEventHandler = New TableManager.EventHandler
    pEventHandler.TextObj Txt
    Set pEventHandler.TableObj = Tbl
    Set pEventHandler.FormObj = Frm
    pEventHandler.Name = Txt.Name
    Set pEventHandler.TextEvent = Txt
End Sub

Public Property Get CmbObj()
    Set CmbObj = pCmbObj
End Property

Public Property Get EvHand()
    Set EvHand = pEventHandler
End Property

Public Property Get Field() As TableManager.CellClass: Set Field = pField: End Property

Public Property Set Field(ByVal Vbl As TableManager.CellClass): Set pField = Vbl: End Property

Public Property Get Name() As String: Name = pName: End Property

Public Property Let Name(ByVal Vbl As String): pName = Vbl: End Property

Private Sub Class_Terminate()
    Set TxtCtl = Nothing
End Sub

