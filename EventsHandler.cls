VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "EventsHandler."

Private LocalEvts As Scripting.Dictionary

Private Sub Class_Initialize(): Set LocalEvts = New Scripting.Dictionary: End Sub

Public Property Get Count() As Integer: Count = LocalEvts.Count: End Property

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
    If TableManager.TableExists(vIndex, Module_Name) Then
        Set Item = LocalEvts(vIndex)
    Else
        LocalEvts.Name = ""
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

Public Sub Add(ByVal Evt As TableManager.TableClass)
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
    
'    Dim Obj As WorksheetClass

'   Error Handling Initialization
    On Error GoTo ErrHandler
           
'   Procedure
    If TableManager.TableExists(Evt.Name, Module_Name) Then Remove Evt.Name
        
    If Evt.Name <> "" Then
        LocalEvts.Add Evt.Name, Evt
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
    Exists = LocalEvts.Exists(vIndex)

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

Public Function ListEvts() As Variant: ListEvts = LocalEvts.Keys: End Function

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
        Set LocalEvts = Nothing
        Set LocalEvts = New Scripting.Dictionary
    Else
        If Not LocalEvts.Exists(vIndex) Then Err.Raise 9
        LocalEvts.Remove vIndex
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


