Attribute VB_Name = "XLAM_Module"
Option Explicit

Private Const Module_Name = "XLAM_Module."

Public Function BuildTable( _
    ByVal WS As Worksheet, _
    ByVal TableName As String _
    ) As Boolean

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
'   09/09/17 RRD Changed to ignore tables with no editable data

'   Declarations
    Const Routine_Name = Module_Name & "BuildTable"
    Dim Tbl As Variant
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    BuildTable = TableManager.Failure
    
'   Procedure

'   Gather the table data
    Set Tbl = New TableManager.TableClass
    If Tbl.CollectData(WS, TableName) Then
        Set Tbl.Form = New TableManager.FormClass
        Tbl.Form.Name = TableName
        
        Tbl.Form.BuildForm (Tbl)
        TableManager.TableAdd Tbl, Module_Name
    End If
    
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

End Function ' BuildTable

Public Sub AutoOpen(ByVal WkBkName As String)
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
    Const Routine_Name = Module_Name & "." & "AutoOpen"
    
    Dim Sht As Worksheet
    Dim Tbl As ListObject
    Dim UserFrm As Object
    Dim SheetClass As TableManager.WorksheetClass
    Dim WkBk As Workbook

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
    Set WkBk = Workbooks(WkBkName)
    TableManager.CheckForVBAProjectAccessEnabled (WkBk.Name)
    
'   Delete existing forms (used for cleanup while debugging)
    For Each UserFrm In ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm Then
            ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
    
'   Procedure
    TableManager.TableSetNewClass Module_Name
    TableManager.WorksheetSetNewClass Module_Name
    
    For Each Sht In WkBk.Worksheets
        Set SheetClass = TableManager.NewWorksheetClass
        Set SheetClass.WS = Sht
        SheetClass.Name = Sht.Name
        TableManager.WorksheetAdd SheetClass, Module_Name
        
        For Each Tbl In Sht.ListObjects
            TableManager.BuildTable Sht, Tbl.Name
        Next Tbl
    Next Sht
    
    DoEvents

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

End Sub      ' AutoOpen


