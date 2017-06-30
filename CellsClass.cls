VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "CellsClass."

Private AllCells As Collection

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
    Dim OneCell As New CellClass   ' Cell Properties
    Const Routine_Name = Module_Name & "Get Item"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    If Me.Exists(vIndex) Then
        Set Item = AllCells(vIndex)
    Else
        AllCells.Name = ""
        Set Item = OneCell
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
    Dim C As CellClass

'   Error Handling Initialization
    On Error GoTo ErrHandler
    Exists = False
    
    Set C = AllCells(vIndex)
    Exists = True
    
ErrHandler:
    Select Case Err.Number
        Case Is = NoError   'Do nothing
        Case Is = 5:        'Does not exist
        Case Is = 13:       'Does not exist
        Case Else:          'Display Error Message Box
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Function

Public Sub Add(ByVal Cll As CellClass)

'   Description:Add a Cell
'   Inputs:     Cll        Cell under analysis
'   Outputs:    *None
'   Requisites:
'   SharedRoutines
'   Example:    AllCells.Add [Cll].ListObject

'     Date
'   05/14/2017 RRD Ini Modification

'   Declarations
    Const Routine_Name = Module_Name & "Add"

'   Error Handling Initialization
    On Error GoTo ErrHandler
           
'   Procedure
    If Exists(Cll) Then Remove Cll
        
    If Cll.HeaderText <> "" Then AllCells.Add Cll
        
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
        Set AllCells = Nothing
        Set AllCells = New Collection
    Else
        If Not AllCells.Exists(vIndex) Then Err.Raise 9
        AllCells.Remove vIndex
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
     Set AllCells = New Collection
End Sub

'   Custom Properties

Public Property Get Count() As Integer
    Count = AllCells.Count
End Property


'   Methods

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemID = -4
'   Description:Enables FOR EACH
    Set NewEnum = AllCells.[_NewEnum]
End Function



