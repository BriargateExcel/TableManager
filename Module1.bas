Attribute VB_Name = "Module1"
Option Explicit

Const Module_Name = "Module1."

Private Frm As FormClass
Private ShtClass As WorksheetClass
Private Tbl As Variant


Private Sub Auto_Open()

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
'   2017-06-17 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "." & "Auto_Open"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    Set ShtClass = New WorksheetClass
    Set ShtClass.Worksheet = ActiveSheet

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

End Sub      ' Auto_Open

Public Sub DeleteForm()
    Dim UserFrm As Object
    For Each UserFrm In ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm Then
            ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
End Sub
    
Public Function BuildTable(ByVal WS As Worksheet, _
    ByVal TableName As String, _
    Target As Range) As Boolean

'   Description: Build a data form for the table
'   Inputs:
'   Target       The cell the user selected
'   TableName   The name of the table containing Target
'   Outputs:
'   Me       Success/Failure
'   Requisites:
'   SharedRoutines
'   Notes:
'   Any notes
'   Example:
'   How to call this routine
'   History
'   06/09/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "BuildTable"
'    Dim UserFrm As Object
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    BuildTable = Failure
    
'   Procedure

'   Delete existing forms (used for cleanup while debugging)
'    For Each UserFrm In ThisWorkbook.VBProject.VBComponents
'        If UserFrm.Type = vbext_ct_MSForm Then
'            ThisWorkbook.VBProject.VBComponents.Remove UserFrm
'        End If
'    Next UserFrm
    
'   Gather the table data
    Set Tbl = New TableClass
    Tbl.CollectData WS, TableName
    
'   Build the form from the table data
    Set Frm = New FormClass
    Frm.BuildForm Tbl, Target
    
    Frm.Show

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

End Function ' BuildTable

Public Function ActiveForm() As FormClass
    Set ActiveForm = Frm
End Function

Public Function ActiveTable() As TableClass
    Set ActiveTable = Tbl
End Function


