Attribute VB_Name = "SharedRoutines"
Option Explicit

Const Module_Name = "SharedRoutines."

Global Const Success = True
Global Const Failure = False
Global Const NoError = 0

Function ActiveCellTableName() As String
'   Function returns table name if active cell is in a table and
'   "" if it isn't.
'    Dim rngActiveCell As Range
    
'    Set rngActiveCell = ActiveCell
'   Statement produces error when active cell is not in a table.
    On Error Resume Next
    ActiveCellTableName = ""
    ActiveCellTableName = ActiveCell.ListObject.Name
    On Error GoTo 0
End Function

Public Function CheckForVBAProjectAccessEnabled() As Boolean

'   Description:
'   Checks that access to the VBA project is enabled
'   If not enabled, tells the user how to enable it
'   Inputs:
'   None
'   Outputs:
'   Me       Success/Failure
'   Requisites:
'   SharedRoutines
'   Notes:
'   Any notes
'   Example:
'   How to call this routine
'   History
'   05/14/2017 RRD Initial Programming

'   Declarations
    Dim VBP As Object ' as VBProject

'   Error Handling Initialization
    On Error GoTo ErrHandler
    CheckForVBAProjectAccessEnabled = Failure

'   Procedure
    If Val(Application.Version) >= 10 Then
        Set VBP = ActiveWorkbook.VBProject
    Else
        MsgBox "This application must be run on Excel 2002 or greater", _
            vbCritical, "Excel Version Check"
        GoTo ErrHandler
    End If

    CheckForVBAProjectAccessEnabled = Success

ErrHandler:
    Set VBP = Nothing

    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Else:
            MsgBox "Your security settings do not allow this procedure to run." _
                & vbCrLf & vbCrLf & "To change your security setting:" _
                & vbCrLf & vbCrLf & " 1. Select Tools - Macro - Security." & vbCrLf _
                & " 2. Click the 'Trusted Sources' tab" & vbCrLf _
                & " 3. Place a checkmark next to 'Trust access to Visual Basic Project.'", _
                vbCritical
    End Select

End Function ' CheckForVBAProjectAccessEnabled

Public Function DspErrMsg(ByVal sRoutine As String)

    Const bDebugMode    As Boolean = True   'Set to false when put into production

    DspErrMsg = MsgBox(Err.Number & ":" & Err.Description, _
                       IIf(bDebugMode, vbAbortRetryIgnore, vbCritical) + _
                         IIf(Err.Number = 999, 0, vbMsgBoxHelpButton), _
                       sRoutine, _
                       Err.HelpFile, _
                       Err.HelpContext)
End Function

