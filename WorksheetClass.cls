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

Private Const Module_Name           As String = "WorksheetClass."

'   Public Properties
Public WithEvents ws As Worksheet
Attribute ws.VB_VarHelpID = -1

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
    If WorksheetExists(Sht.Name, Module_Name) Then WorksheetRemove Sht.Name, Module_Name
        
    If Sht.Name <> "" Then
        Set Obj = New EventHandler
        WorksheetAdd Sht, Module_Name
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
Dim TableName As String
Dim Tbl As TableClass
    
    Cancel = True

    TableName = ActiveCellTableName
    
    If TableName <> "" Then
        Set Tbl = TableItem(TableName, Module_Name)
        Set Tbl.Target = Target
        Tbl.Form.PopulateForm Tbl
        Tbl.Form.Show
    Else
        MsgBox "The selected cell is not inside a table. " & _
            "Double-click a cell inside a table.", _
            vbOKOnly Or vbExclamation, _
            "Selected Cell Not in a Table"
    End If
    
End Sub

Public Sub CollectData(ByVal ws As Worksheet)
    pName = ws.Name
End Sub

Public Property Get Name() As String: Name = pName: End Property

Public Property Let Name(ByVal Vbl As String): pName = Vbl: End Property

Private Sub Class_Terminate()
    Set ws = Nothing
End Sub
