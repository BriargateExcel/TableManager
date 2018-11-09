Attribute VB_Name = "XLAM_Module"
'@Folder("TableManager.Main")

Option Explicit

Private Const Module_Name As String = "XLAM_Module."

Private Init As Boolean
Private pMainWorkbook As Workbook

Private LastControl As control

' TODO Implement more specific error messages
Public Enum CustomError

    Success = 0

    [_First] = vbObjectError - 10000

    ArrayMustBe1or2Dimensions                    ' description

    CustomErrorTwo                               ' description

    ' ... more error names

    [_Last]

End Enum

Public Sub SetLastControl(ByVal Ctl As control)
    Set LastControl = Ctl
End Sub

Public Function GetLastControl() As control
    Set GetLastControl = LastControl
End Function

Public Function GetMainWorkbook() As Workbook
    Set GetMainWorkbook = pMainWorkbook
End Function

Public Sub SetMainWorkbook(ByVal Wkbk As Workbook)
    Set pMainWorkbook = Wkbk
End Sub

Public Function GetWorkBookPath(ByVal Wkbk As Workbook) As String
    GetWorkBookPath = Wkbk.Path
End Function

Public Sub SetUpWorkbook(ByVal Wkbk As Workbook)
    
    Dim Sht As Worksheet
    Dim TblObj As ListObject
    Dim UserFrm As Object
    Dim WkSht As WorksheetClass
    
    Const RoutineName As String = Module_Name & "AutoOpen"
    On Error GoTo ErrorHandler
    
    SetInitializing
    
    Set pMainWorkbook = Wkbk
    
    If Not CheckForVBAProjectAccessEnabled(Wkbk) Then
        MsgBox "You must set the project access for the " & _
               "TableManager Add-In to work", _
               vbOKOnly Or vbCritical, _
               "Project Access"
    
    End If
    
    ' Delete all the old UserForms
    If Wkbk.Name = ThisWorkbook.Name Then
        For Each UserFrm In Wkbk.VBProject.VBComponents
            If UserFrm.Type = vbext_ct_MSForm And _
               Left$(UserFrm.Name, 8) = "UserForm" _
               Then
                Application.Wkbk.VBProject.VBComponents.Remove UserFrm
            End If
        Next UserFrm
    End If
    
    WorksheetSetNewClass Module_Name
    TableSetNewClass Module_Name
    
    ' Go through all the worksheets and all the tables on each worksheet
    ' collecting the data and building the form for each table
    For Each Sht In Wkbk.Worksheets
        Set WkSht = New WorksheetClass
        Set WkSht.Worksheet = Sht
        WkSht.Name = Sht.Name
        
        WorksheetAdd WkSht, Module_Name
        
        For Each TblObj In Sht.ListObjects
            BuildTable Wkbk, TblObj, Module_Name
        Next TblObj
    
    Next Sht
    
    DoEvents
    
    ReSetInitializing

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


