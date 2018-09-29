VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Events")

Option Explicit

Private Const Module_Name As String = "EventClass."

Private WithEvents pSheetEvent As Worksheet
Attribute pSheetEvent.VB_VarHelpID = -1

Private ptableobj As TableClass

Private WithEvents pFormEvent As MSForms.UserForm
Attribute pFormEvent.VB_VarHelpID = -1
Private pFormObj As Object

Private WithEvents pButtonEvent As MSForms.CommandButton
Attribute pButtonEvent.VB_VarHelpID = -1
Private pButtonObj As MSForms.CommandButton

Private WithEvents pTextEvent As MSForms.TextBox
Attribute pTextEvent.VB_VarHelpID = -1
Private pTextObj As MSForms.TextBox

Private WithEvents pComboEvent As MSForms.ComboBox
Attribute pComboEvent.VB_VarHelpID = -1
Private pComboObj As MSForms.ComboBox

Private pName As String

Private LastEvent As String

Public Property Set SheetEvent(ByVal Vbl As Worksheet): Set pSheetEvent = Vbl: End Property

Public Property Set FormEvent(ByVal Vbl As MSForms.UserForm): Set pFormEvent = Vbl: End Property

Public Property Set ButtonEvent(ByVal Vbl As MSForms.CommandButton): Set pButtonEvent = Vbl: End Property

Public Property Set TextEvent(ByVal Vbl As MSForms.TextBox): Set pTextEvent = Vbl: End Property

Public Property Set ComboEvent(ByVal Vbl As MSForms.ComboBox): Set pComboEvent = Vbl: End Property

Public Property Set TableObj(ByVal Tbl As TableClass)
    Debug.Assert Initializing

    Set ptableobj = Tbl
End Property

Public Property Set FormObj(ByVal FO As Object)
    Set pFormObj = FO
    Set pFormEvent = FO
End Property

Public Property Set ButtonObj(ByVal Btn As MSForms.CommandButton)
    Set pButtonObj = Btn
    Set pButtonEvent = Btn
End Property                                     ' ButtonObj

Public Property Set TextObj(ByVal Txt As MSForms.TextBox)
    Set pTextObj = Txt
    Set pTextEvent = Txt
End Property                                     ' TextObj

Public Property Set ComboObj(ByVal Cmb As MSForms.ComboBox)
    Debug.Assert Initializing

    Set pComboEvent = Cmb
    Set pComboObj = Cmb
End Property                                     ' ComboObj

Private Sub pFormEvent_MouseMove( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    Const RoutineName As String = Module_Name & "FormEvent_MouseMove"
    On Error GoTo ErrorHandler
    
    If LastEvent = "Form Event" Then Exit Sub
    LastEvent = "Form Event"
    
    LowLightControl GetLastControl, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    Debug.Print "mousemove error number: " & Err.Number
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' FormEvent_MouseMove

Private Sub pButtonEvent_MouseMove( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)
    
    Const RoutineName As String = Module_Name & "ButtonEvent_MouseMove"
    On Error GoTo ErrorHandler
    
    If ptableobj Is Nothing Then
        ' Processing the database form
        '        Debug.Print "db"
    Else
        ' Processing a table form
        If LastEvent = "Button Event" Then Exit Sub
        LastEvent = "Button Event"
        
        ' Exit if button already properly set
        ' Reduces the flickering of the user form
        If pButtonObj.BackColor = LightestColorValue Then Exit Sub
        
        PaintButtons
    End If
    
    LowLightControl GetLastControl, Module_Name
    SetLastControl pButtonObj
    
    HighLightControl pButtonObj, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' ButtonEvent_MouseMove

Private Sub pTextEvent_MouseMove( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    Const RoutineName As String = Module_Name & "TextEvent_MouseMove"
    On Error GoTo ErrorHandler
    
    LowLightControl GetLastControl, Module_Name
    SetLastControl pTextObj
    
    ' Highlight the selected Text Box
    HighLightControl pTextObj, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' TextEvent_MouseMove

Private Sub pComboEvent_MouseMove( _
        ByVal Combo As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    Const RoutineName As String = Module_Name & "ComboEvent_MouseMove"
    On Error GoTo ErrorHandler
    
    LowLightControl GetLastControl, Module_Name
    SetLastControl pComboObj
    
    HighLightControl pComboObj, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' ComboEvent_MouseMove

Private Sub pSheetEvent_BeforeDoubleClick( _
        ByVal Target As Range, _
        Cancel As Boolean)
    
    Const RoutineName As String = Module_Name & "SheetEvent_BeforeDoubleClick"
    On Error GoTo ErrorHandler
    
    Dim TableName As String
    Dim Tbl As TableClass
    
    Cancel = True

    TableName = ActiveCellTableName
    If TableName = vbNullString Then
        MsgBox "Please select a cell in the body of the table", _
               vbOKOnly Or vbExclamation, "Select a Table Cell"
        Exit Sub
    End If
    
    Set Tbl = New TableClass
    Set Tbl = TableItem(TableName, Module_Name)
    Set Tbl.ActiveTarget = Target
    If Tbl Is Nothing Then  ' Means the table has no UserForm
        MsgBox "The user cannot edit this table", vbOKOnly Or vbCritical, "Table Not User-Editable"
        Exit Sub
    End If

    Set Tbl.Table = Tbl.ActiveTarget.ListObject
    
    Dim Isect As Range
    
    Set Isect = Application.Intersect(Target, Tbl.Table.HeaderRowRange)
    If Not Isect Is Nothing Then
        BuildDataBaseForm Tbl, Module_Name
        ShowAnyForm DataBaseFormName
        Exit Sub
    End If
    
    PopulateForm Tbl, Module_Name
    ShowAnyForm Tbl.Form.Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
End Sub                                          ' SheetEvent_BeforeDoubleClick

Private Sub pButtonEvent_Click()
    Dim Target As Range
    Dim TableName As String
    Dim Tbl As TableClass
    
    Const RoutineName As String = Module_Name & "ButtonEvent_Click"
    On Error GoTo ErrorHandler
    
    TableName = ActiveCellTableName
    Set Tbl = New TableClass
    Set Tbl = Table(TableName, Module_Name)
    
    Select Case pButtonEvent.Caption
        ' Table Form Buttons
    Case "Validate Form"
        If ValidateForm(Tbl, Module_Name) Then
            MsgBox "There are no validation errors", _
                   vbOKOnly, "Successful Copy"
        Else
            MsgBox "There is/are a validation error(s)", _
                   vbOKOnly Or vbExclamation, _
                   "Data Validation Error"
        End If
    
    Case "Copy To Table"
        If ValidateForm(Tbl, Module_Name) Then
            PopulateTable Tbl, Module_Name
            MsgBox "Form data successfully copied to Table row", _
                   vbOKOnly, "Successful Copy"
        Else
            MsgBox "There is/are a validation error(s). " & _
                   "Correct the error then copy the Form to the Table row.", _
                   vbOKOnly Or vbExclamation, _
                   "Data Validation Error"
        End If
    
    Case "Next Row"
        TurnOnCellDescriptions Tbl, Module_Name
        ptableobj.NextRow
        Set ptableobj.ActiveTarget = ptableobj.ActiveTarget.Offset(0, 0)
        PopulateForm Tbl, Module_Name
    
    Case "Previous Row"
        TurnOnCellDescriptions Tbl, Module_Name
        ptableobj.PreviousRow
        Set ptableobj.ActiveTarget = ptableobj.ActiveTarget.Offset(0, 0)
        PopulateForm Tbl, Module_Name
    
    Case "Insert Above"
        ptableobj.ActiveTarget.ListObject.ListRows.Add ptableobj.ActiveRow
        Set ptableobj.ActiveTarget = ptableobj.ActiveTarget.Offset(-1, 0)
        ClearForm Tbl, Module_Name
    
    Case "Insert Below"
        ptableobj.ActiveTarget.ListObject.ListRows.Add ptableobj.ActiveRow + 1
        Set ptableobj.ActiveTarget = ptableobj.ActiveTarget.Offset(1, 0)
        ClearForm Tbl, Module_Name
    
    Case "Insert at Top"
        ptableobj.ActiveTarget.ListObject.ListRows.Add 1
        Set ptableobj.ActiveTarget = ptableobj.FirstCell
        ClearForm Tbl, Module_Name
    
    Case "Insert at Bottom"
        ptableobj.ActiveTarget.ListObject.ListRows.Add
        Set ptableobj.ActiveTarget = ptableobj.LastCell
        ClearForm Tbl, Module_Name
    
    Case "Delete This Row"
        If ptableobj.LastRow = ptableobj.ActiveRow Then
            Set Target = ptableobj.ActiveTarget.Offset(-1, 0)
        Else
            Set Target = ptableobj.ActiveTarget.Offset(1, 0)
        End If
        ptableobj.ActiveTarget.ListObject.ListRows(ptableobj.ActiveRow).Delete
        Set ptableobj.ActiveTarget = Target
        PopulateForm Tbl, Module_Name
    
    Case "Go To Top"
        TurnOnCellDescriptions ptableobj, Module_Name
        Set ptableobj.ActiveTarget = ptableobj.FirstCell
        TurnOffCellDescriptions ptableobj, Module_Name
        PopulateForm Tbl, Module_Name
    
    Case "Go To Bottom"
        TurnOnCellDescriptions ptableobj, Module_Name
        Set ptableobj.ActiveTarget = ptableobj.LastCell
        TurnOffCellDescriptions ptableobj, Module_Name
        PopulateForm Tbl, Module_Name
    
    Case Is = "Cancel"
        TurnOnCellDescriptions ptableobj, Module_Name
        pFormObj.Hide
        Exit Sub
    
        ' Start of DataBase Form Buttons
    Case Is = "Copy to File"
        OutputTable Module_Name
        Exit Sub
    
    Case Is = "Fetch From File"
        InputTable Module_Name
        Exit Sub
    
    Case Is = "Change File"
        ChangeFile ptableobj, Module_Name
        Exit Sub
    
    Case Is = "Cancel File Processing"
        pFormObj.Hide
        Exit Sub
        ' End of DataBase Form Buttons

    Case Else
        MsgBox "This button does not exist: " & _
               pButtonEvent.Caption & ". Select another.", _
               vbOKOnly Or vbExclamation, "Button Does Not Exist"

    End Select
    
    PaintButtons
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

    '    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' ButtonEvent_Click

Private Sub PaintButtons()
    '   Sets all the button colors to reflect the current state of the form
    ' Assumes that Next and Previous buttons are named properly

    Const RoutineName As String = Module_Name & "PaintButtons"
    On Error GoTo ErrorHandler
    
    Debug.Assert Not Initializing
    
    Dim Ctl As Variant
    
    For Each Ctl In pFormObj.Controls
        If Left$(Ctl.Name, 7) = "Command" Then   ' All buttons names start with "CommandButton"
            EnableButton Ctl, Module_Name
            LowLightControl Ctl, Module_Name
                
            If Ctl.Caption = "Previous Row" And ptableobj.ActiveRow = 1 Then
                DisableButton Ctl, Module_Name
            End If
                
            If Ctl.Caption = "Next Row" And ptableobj.ActiveRow = ptableobj.LastRow Then
                DisableButton Ctl, Module_Name
            End If
                
        End If
    Next Ctl
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' PaintButtons

Public Property Get Name() As String: Name = pName: End Property
Public Property Let Name(ByVal Vbl As String): pName = Vbl: End Property


