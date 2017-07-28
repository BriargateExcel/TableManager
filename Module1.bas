Attribute VB_Name = "Module1"
Option Explicit

Const Module_Name = "Module1."

Public AllTbls As TablesClass ' ToDo: does this need to be public? Is there some other way to do this?
Public AllShts As WorksheetsClass ' ToDo: does this need to be public? Is there some other way to do this?

Public Const DarkestColor = &H763232 ' AF Dark Blue
Public Const LightestColor = &HE7E2E2 ' AF Light Gray

Public Const LabelBackGround = DarkestColor
Public Const LabelFont = LightestColor

Public Const ButtonNothingBackGround = DarkestColor
Public Const ButtonNothingFont = LightestColor

Public Const ButtonHighLightBackGround = LightestColor
Public Const ButtonHighLightFont = DarkestColor


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
    
    Dim Sht As Worksheet
    Dim ThisSheet As Object
    Dim Tbl As ListObject
    Dim UserFrm As Object
    Dim SheetClass As WorksheetClass

'   Error Handling Initialization
    On Error GoTo ErrHandler
    CheckForVBAProjectAccessEnabled
    
'   Delete existing forms (used for cleanup while debugging)
    For Each UserFrm In ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm Then
            ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
    
'   Procedure
    Set AllTbls = New TablesClass
    Set AllShts = New WorksheetsClass
    
    For Each Sht In ThisWorkbook.Worksheets
        For Each Tbl In Sht.ListObjects
            BuildTable Sht, Tbl.Name
            Set SheetClass = New WorksheetClass
            Set SheetClass.WS = Sht
            SheetClass.Name = Sht.Name
            AllShts.Add SheetClass
        Next Tbl
    Next Sht
    
    DoEvents

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

'   Declarations
    Const Routine_Name = Module_Name & "BuildTable"
    Dim Tbl As Variant
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    BuildTable = Failure
    
'   Procedure

'   Gather the table data
    Set Tbl = New TableClass
    Tbl.CollectData WS, TableName
    Set Tbl.Form = New FormClass
    Tbl.Form.Name = TableName
    
    Tbl.Form.BuildForm (Tbl)
'    Tbl.Add Tbls(TableName)
    AllTbls.Add Tbl
    
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



