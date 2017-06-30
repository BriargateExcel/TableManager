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

Private Const Module_Name = "CtlsClass."

Private AllCtls As Scripting.Dictionary

Public Property Get Item(ByVal vIndex As Variant)
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
'   Description:Get a Field
'   Inputs:     vIndex      Name or Number
'   Outputs:    Me          Header Object
'   Requisites: Function:   me.Add
'   Example:    ? Header.Item(sName).Fields.Count
'               ? Header.Item(1).Fields.Count
'               ? Header(1).Fields.Count

'     Date   Ini Modification
'   05/14/2017 RRD Initial Programming

'   Declarations
    Dim OneCtl As New ControlClass   ' Control Properties
    Const Routine_Name = Module_Name & "Get Item"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    If Me.Exists(vIndex) Then
        Set Item = AllCtls(vIndex)
    Else
        AllCtls.Name = ""
        Set Item = OneCtl
    End If

ErrHandler:
    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Is = 5: Set Item = Nothing
        Case Else:          'Display Error Message Box
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Property

Public Function Exists(ByVal vIndex) As Boolean

'   Description:Determine if OneCell exists in AllCells
'   Inputs:     vIndex      Name or Instance #
'   Outputs:    Me          True/False
'   Requisites: *None
'   Example:    ? AllCells.Exists("DET")
'               ? AllCells.Exists(1)

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Exists"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    Exists = AllCtls.Exists(vIndex)
    
ErrHandler:
    Select Case Err.Number
        Case Is = NoError   'Do nothing
        Case Is = 424:        'Does not exist
        Case Is = 13:       'Does not exist
        Case Else:          'Display Error Message Box
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Function

Public Sub Add(ByVal Ctl As Control)

'   Description:Add a control
'   Inputs:     Ctl The control to be added
'   Outputs:    *None
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
        Set Obj = New EventHandler
        AllCtls.Add Ctl.Name, Obj
    End If
        
ErrHandler:
    Select Case Err.Number
        Case Is = NoError:   'Do nothing
        Case Else:          'Display Error Message Box
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select
    
End Sub


Public Sub Remove(ByVal vIndex As Variant)

'   Description:Close and Remove a connection
'   Inputs:     Name        Connection name
'   Outputs:    *None
'   Requisites: *None
'   Example:    AllCells(1).Add "BXL"

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Remove"

'   Error Handling Initialization
    On Error GoTo ErrHandler
   
'   Procedure
    If CStr(vIndex) = "*" Then
        Set AllCtls = Nothing
        Set AllCtls = New Collection
    Else
        If Not AllCtls.Exists(vIndex) Then Err.Raise 9
        AllCtls.Remove vIndex
    End If

ErrHandler:
    Select Case Err.Number
        Case Is = NoError:   'Do nothing
        Case Else:          'Display Error Message Box
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select
    
End Sub


'   Internal Routines

Private Sub Class_Initialize()
     Set AllCtls = New Scripting.Dictionary
End Sub

'   Custom Properties

Public Property Get Count() As Integer
    Count = AllCtls.Count
End Property

Public Function ListCtls() As Variant

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
    Const Routine_Name = Module_Name & "." & "ListCtls"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    ListCtls = AllCtls.Keys

    ListCtls = Success
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

End Function ' ListCtls
