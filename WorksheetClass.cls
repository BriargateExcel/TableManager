VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const Module_Name           As String = "WorksheetClass"

'   Public Properties
Public WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1

Private pName As String

Public Sub Add(ByVal Sht As Worksheet)

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
    If AllShts.Exists(Sht.Name) Then AllShts.Remove Sht.Name
        
    If Sht.Name <> "" Then
        Set Obj = New EventHandler
        AllShts.Add Sht
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

Private Sub WS_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim Rng As Range
Dim TableName As String
Dim WS As Worksheet
Dim Tbl As TableClass
    
    Cancel = True

    TableName = ActiveCellTableName
    
    If TableName <> "" Then
        Set Tbl = AllTbls(TableName)
        Set Tbl.Target = Target
        Tbl.Form.PopulateForm Tbl, Target
        Tbl.Form.Show
    End If
    
End Sub

Public Sub CollectData(ByVal WS As Worksheet)
    pName = WS.Name
End Sub

Public Property Get Name() As String: Name = pName: End Property

Public Property Let Name(ByVal Vbl As String): pName = Vbl: End Property

Private Sub Class_Terminate()
    Set WS = Nothing
End Sub
