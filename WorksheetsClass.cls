VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "WorksheetsClass."

Private LocalShts As Scripting.Dictionary

Private Sub Class_Initialize(): Set LocalShts = New Scripting.Dictionary: End Sub

Public Property Get Count() As Integer: Count = TableManager.TableCount(Module_Name): End Property

Public Property Get Item(ByVal vIndex As Variant)
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

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    If Me.Exists(vIndex) Then
        Set Item = TableManager.TableItem(vIndex, Module_Name)
    Else
        Set Item = New TableManager.TableClass
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

Public Sub Add(ByVal Sht As TableManager.WorksheetClass)
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
    If LocalShts.Exists(Sht.Name) Then Remove Sht.Name
        
    If Sht.Name <> "" Then ' The sheet must have a name
        LocalShts.Add Sht.Name, Sht
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
    
End Sub ' Add

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
    Exists = TableManager.TableExists(vIndex, Module_Name)
    
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

Public Function ListTbls() As Variant
'   Description: Description of what function does
'   Inputs:
'   Outputs:
'   Me       Success/Failure
'   Requisites:
'   None
'   Notes:
'   Any notes
'   Example:
'   How to call this routine
'   History
'   2017-06-13 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "." & "ListTbls"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    ListTbls = LocalShts.Keys

    ListTbls = TableManager.Success
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Else:
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select

End Function ' ListTbls

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
        TableManager.TableSetNothing Module_Name
        TableManager.TableSetNewDict Module_Name
    Else
        If Not TableManager.TableExists(vIndex, Module_Name) Then Err.Raise 9
        TableManager.TableRemove vIndex, Module_Name
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

