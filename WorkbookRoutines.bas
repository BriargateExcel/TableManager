Attribute VB_Name = "WorkbookRoutines"
'@Folder("TableManager.Workbooks")

Option Explicit

Private Const Module_Name As String = "WorkbookRoutines."

Private pAllBooks As WorkbooksClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TableRoutines.")
End Function                                     ' ModuleList

Public Sub AddTableToWorkbook(ByVal Wkbk As WorkbookClass)
    pAllBooks.Add Wkbk, Module_Name
End Sub

Public Function GetWorkbook(ByVal WorkbookName As String) As WorkbookClass
    Dim TempWorkbook As WorkbookClass
    Set TempWorkbook = New WorkbookClass
    Set TempWorkbook = pAllBooks.Item(WorkbookName, Module_Name)
    Set GetWorkbook = TempWorkbook
End Function

Public Sub WorkbookAdd( _
       ByVal WS As Variant, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "WorkbookAdd"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    pAllBooks.Add WS, ModuleName

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorkbookAdd

Public Sub WorkbookSetNewClass(ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "WorkbookSetNewClass"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllBooks = New WorkbooksClass

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorkbookSetNewClass


