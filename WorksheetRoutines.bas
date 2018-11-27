Attribute VB_Name = "WorksheetRoutines"
'@Folder("TableManager.Worksheets")

Option Explicit

Private Const Module_Name As String = "WorksheetRoutines."

Private pAllShts As WorksheetsClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TableRoutines.")
End Function                                     ' ModuleList

Public Sub WorksheetAdd( _
       ByVal WS As Variant, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "WorksheetAdd"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    pAllShts.Add WS, ModuleName

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorksheetAdd

Public Sub WorksheetSetNewClass(ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "WorksheetSetNewClass"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllShts = New WorksheetsClass

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorksheetSetNewClass


