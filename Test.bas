Attribute VB_Name = "test"
Option Explicit

Const Module_Name As String = "test."

Public Sub test()
UnifyDataValidation TableManager.TableItem("ReqTable", Module_Name)
End Sub

Public Sub UnifyDataValidation(ByVal Tbl As TableManager.TableClass)
    
    Const RoutineName As String = Module_Name & "UnifyDataValidation"
    On Error GoTo ErrorHandler
    
    Debug.Assert Not Initializing
    
    Dim I As Long
    Dim J As Long
    
    Dim Rng As Range
    
    Set Rng = Tbl.DBRange
    
    For I = 2 To Tbl.NumRows
        For J = 1 To Tbl.CellCount
            If HasVal(Rng(I, J)) Then
'                If Rng(I, J).Validation.Type <> Rng(1, J).Validation.Type Then Stop
'                If Rng(I, J).Validation.IgnoreBlank <> Rng(1, J).Validation.IgnoreBlank Then Stop
'                If Rng(I, J).Validation.AlertStyle <> Rng(1, J).Validation.AlertStyle Then Stop
'                If Rng(I, J).Validation.Operator <> Rng(1, J).Validation.Operator Then Stop
'                If Rng(I, J).Validation.ShowInput <> Rng(1, J).Validation.ShowInput Then Stop
'                If Rng(I, J).Validation.InputTitle <> Rng(1, J).Validation.InputTitle Then Stop
'                If Rng(I, J).Validation.InputMessage <> Rng(1, J).Validation.InputMessage Then Stop
'                If Rng(I, J).Validation.ShowError <> Rng(1, J).Validation.ShowError Then Stop
'                If Rng(I, J).Validation.ErrorTitle <> Rng(1, J).Validation.ErrorTitle Then Stop
'                On Error Resume Next
'                If Rng(I, J).Validation.Formula1 <> Rng(1, J).Validation.Formula1 Then Stop
'                If Rng(I, J).Validation.Formula2 <> Rng(1, J).Validation.Formula2 Then Stop
'                On Error GoTo ErrorHandler
            End If
        Next J
    Next I
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

