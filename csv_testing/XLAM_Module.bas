Attribute VB_Name = "XLAM_Module"
Option Explicit

Private Const Module_Name As String = "XLAM_Module."

Private Init As Boolean
Private pMainWorkbook As Workbook

Private LastControl As Control

' TODO Implement more specific error messages
Public Enum CustomError
' https://www.linkedin.com/pulse/interesting-technique-error-handling-enumeration-vba-chip-pearson/

    Success = 0

    [_First] = vbObjectError - 10000

    ArrayMustBe1or2Dimensions ' description

    CustomErrorTwo ' description

    ' ... more error names

    [_Last]

End Enum

Public Function IsValidErrNum(ByVal ErrNum As CustomError) As Boolean

    IsValidErrNum = (ErrNum = CustomError.Success) Or _
                    ((ErrNum > CustomError.[_First]) And (ErrNum < CustomError.[_Last]))

End Function

Public Sub SetLastControl(ByVal Ctl As Control)
    Set LastControl = Ctl
End Sub

Public Function GetLastControl() As Control
    Set GetLastControl = LastControl
End Function

Public Function GetMainWorkbook() As Workbook
    Set GetMainWorkbook = pMainWorkbook
End Function

Public Sub SetMainWorkbook(ByVal Wkbk As Workbook)
    Set pMainWorkbook = Wkbk
End Sub

Public Sub AutoOpen(ByVal Wkbk As Workbook)
    
    Dim Sht As Worksheet
    Dim Tbl As ListObject
    Dim UserFrm As Object
    Dim WkSht As TableManager.WorksheetClass
    
    Const RoutineName As String = Module_Name & "AutoOpen"
    On Error GoTo ErrorHandler
    
    Init = True
    Set pMainWorkbook = Wkbk
    
    If Not CheckForVBAProjectAccessEnabled(ThisWorkbook) Then
        MsgBox "You must set the project access for the " & _
               "TableManager Add-In to work", _
               vbOKOnly Or vbCritical, _
               "Project Access"
    
    End If
    
    For Each UserFrm In Application.ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm And _
           Left$(UserFrm.Name, 8) = "UserForm" _
           Then
            Application.ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
    
    TableManager.WorksheetSetNewClass Module_Name
    TableManager.TableSetNewClass Module_Name
    
    For Each Sht In GetMainWorkbook.Worksheets
        Set WkSht = New TableManager.WorksheetClass
        Set WkSht.Worksheet = Sht
        WkSht.Name = Sht.Name
        
        For Each Tbl In Sht.ListObjects
            BuildTable WkSht, Tbl, Module_Name
        Next Tbl
        
        TableManager.WorksheetAdd WkSht, Module_Name
    Next Sht
    
    DoEvents
    
    Init = False

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub                                          ' AutoOpen

Public Function Initializing() As Boolean
    Initializing = Init
End Function                                     ' Initializing

Public Sub SetInitializing()
    Init = True
End Sub

Public Sub ReSetInitializing()
    Init = False
End Sub


