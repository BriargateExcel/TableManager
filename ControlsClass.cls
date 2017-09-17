VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "ControlsClass."

Private LocalCtls As Scripting.Dictionary

Private Sub Class_Initialize(): Set LocalCtls = New Scripting.Dictionary: End Sub

Public Property Get Controls(): Controls = LocalCtls.Keys: End Property

Public Property Get Count() As Integer: Count = LocalCtls.Count: End Property

Public Property Get Item(ByVal vIndex As Variant)
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
'   Description:Get a Field
'   Inputs: vIndex      Name or Number
'   Outputs: Me          Header Object
'   Requisites: Function: me.Add
'   Example: ? Header.Item(sName).Fields.Count
'               ? Header.Item(1).Fields.Count
'               ? Header(1).Fields.Count

'     Date   Ini Modification
'   05/14/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "Get Item"
    Dim NewCtl As TableManager.ControlClass

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    If Me.Exists(vIndex) Then
        Set Item = LocalCtls(vIndex)
    Else
        Set NewCtl = New TableManager.ControlClass
        LocalCtls.Add vIndex, NewCtl
    End If

ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Is = 5: Set Item = Nothing
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select

End Property

Public Sub Add(ByVal Ctl As TableManager.ControlClass)
'   Description:Add a control
'   Inputs: Ctl The control to be added
'   Outputs: *None
'   Requisites:
'   SharedRoutines
'   Example:

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Add"
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
           
'   Procedure
    If Exists(Ctl.Name) Then Remove Ctl.Name
        
    If Ctl.Name <> "" Then
        LocalCtls.Add Ctl.Name, Ctl
    End If
        
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select
    
End Sub

Public Sub AddEventHandler(ByVal Ctl As Control)
'   Description:Add a control
'   Inputs: Ctl The control to be added
'   Outputs: *None
'   Requisites:
'   SharedRoutines
'   Example:

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Add"
    
    Dim Obj As Object

'   Error Handling Initialization
    On Error GoTo ErrHandler
           
'   Procedure
    If Exists(Ctl.Name) Then Remove Ctl.Name
        
    If Ctl.Name <> "" Then
        Set Obj = New TableManager.EventHandler
        LocalCtls.Add Ctl.Name & "EH", Obj
    End If
        
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select
    
End Sub

Public Function Exists(ByVal vIndex) As Boolean
'   Description:Determine if OneCell exists in AllCells
'   Inputs: vIndex      Name or Instance #
'   Outputs: Me          True/False
'   Requisites: *None
'   Example: ? AllCells.Exists("DET")
'               ? AllCells.Exists(1)

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Exists"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    Exists = LocalCtls.Exists(vIndex)
    
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError   'Do nothing
        Case Is = 424: 'Does not exist
        Case Is = 13: 'Does not exist
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select

End Function

Public Sub Remove(ByVal vIndex As Variant)
'   Description:Close and Remove a connection
'   Inputs: Name        Connection name
'   Outputs: *None
'   Requisites: *None
'   Example: AllCells(1).Add "BXL"

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Remove"

'   Error Handling Initialization
    On Error GoTo ErrHandler
   
'   Procedure
    If CStr(vIndex) = "*" Then
        Set LocalCtls = Nothing
        Set LocalCtls = New Collection
    Else
        If Not LocalCtls.Exists(vIndex) Then Err.Raise 9
        LocalCtls.Remove vIndex
    End If

ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select
    
End Sub

