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

Public Function NewWorkbookClass() As WorkbookClass
    Set NewWorkbookClass = New WorkbookClass
End Function
    
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

Public Sub InitializeWorkbookForTableManager(ByVal Wkbk As Workbook, _
    Optional ByVal KeepUserForms As Boolean = True)
    
    Const RoutineName As String = Module_Name & "InitializeWorkbookForTableManager"
    On Error GoTo ErrorHandler
    
    SetInitializing
    
    If Not CheckForVBAProjectAccessEnabled(Wkbk) Then
        MsgBox "You must set the project access for the " & _
               "TableManager Add-In to work", _
               vbOKOnly Or vbCritical, _
               "Project Access"
        Stop
    End If
    
    If Not KeepUserForms Then
        Set pMainWorkbook = Wkbk
        
        ' Delete all the old UserForms from TableManager
        ' I haven't found a way to create and add a new userform to another workbook
        Dim UserFrm As Object
        For Each UserFrm In ThisWorkbook.VBProject.VBComponents
            If UserFrm.Type = vbext_ct_MSForm And _
               Left$(UserFrm.Name, 8) = "UserForm" _
               Then
                ThisWorkbook.VBProject.VBComponents.Remove UserFrm
            End If
        Next UserFrm
        WorksheetSetNewClass Module_Name
        TableSetNewClass Module_Name
    End If
    
    
    ' Go through all the worksheets and all the tables on each worksheet
    ' collecting the data and building the form for each table
    Dim WkSht As WorksheetClass
    Dim Sht As Worksheet
    Dim TblObj As ListObject
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
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    '    DisplayError RoutineName

End Sub                                          ' InitializeWorkbookForTableManager

Public Function Initializing() As Boolean
    Initializing = Init
End Function                                     ' Initializing

Public Sub SetInitializing()
    Init = True
End Sub

Public Sub ReSetInitializing()
    Init = False
End Sub


