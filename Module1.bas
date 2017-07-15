Attribute VB_Name = "Module1"
Option Explicit

Const Module_Name = "Module1."

Private frm As FormClass
Private ShtClass As WorksheetClass
Private Tbl As Variant

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
    
Public Function BuildTable( _
    ByVal WS As Worksheet, _
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
    Set frm = New FormClass
    frm.BuildForm Tbl, Target
    
    frm.Show

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
    Set ActiveForm = frm
End Function

Public Function ActiveTable() As TableClass
    Set ActiveTable = Tbl
End Function


Sub colors56()
'   57 colors, 0 to 56
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   'pre XL97 xlManual
    Dim i As Long
    Dim str0 As String, str As String
    
    For i = 0 To 56
        Cells(i + 1, 1).Interior.ColorIndex = i
        Cells(i + 1, 1).Value = "[Color " & i & "]"
        
        Cells(i + 1, 2).Font.ColorIndex = i
        Cells(i + 1, 2).Value = "[Color " & i & "]"
        
        str0 = Right("000000" & Hex(Cells(i + 1, 1).Interior.Color), 6)
'       Excel shows nibbles in reverse order so make it as RGB
        str = Right(str0, 2) & Mid(str0, 3, 2) & Left(str0, 2)
'       generating 2 columns in the HTML table
        Cells(i + 1, 3) = "#" & str & "#" & str & ""
        
        Cells(i + 1, 4).Formula = "=Hex2dec(""" & Right(str0, 2) & """)"
        
        Cells(i + 1, 5).Formula = "=Hex2dec(""" & Mid(str0, 3, 2) & """)"
        
        Cells(i + 1, 6).Formula = "=Hex2dec(""" & Left(str0, 2) & """)"
        
        Cells(i + 1, 7) = "[Color " & i & ")"
    Next i
    
done:
    Application.Calculation = xlCalculationAutomatic  'pre XL97 xlAutomatic
    Application.ScreenUpdating = True
End Sub

