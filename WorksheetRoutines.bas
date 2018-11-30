Attribute VB_Name = "WorksheetRoutines"
'@Folder("TableManager.Worksheets")

Option Explicit

Private Const Module_Name As String = "WorksheetRoutines."

Private pAllShts As WorksheetsClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TableRoutines.", "WorksheetsClass.")
End Function                                     ' ModuleList

Public Sub RemoveWorksheet(ByVal WkShtName As String)
    pAllShts.Remove WkShtName, Module_Name
End Sub

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

Public Sub WorksheetSetNothing(ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "WorksheetSetNothing"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllShts = Nothing

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorksheetSetNothing

Public Sub WorksheetSetNewDict(ByVal ModuleName As String)
    ' Used in WorksheetsClass

    Const RoutineName As String = Module_Name & "WorksheetSetNewDict"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllShts = New Scripting.Dictionary

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorksheetSetNewDict

Public Function WorksheetExists( _
       ByVal Sht As Variant, _
       ByVal ModuleName As String _
       ) As Boolean

    Const RoutineName As String = Module_Name & "WorksheetExists"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    WorksheetExists = pAllShts.Exists(Sht, Module_Name)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' WorksheetExists

Public Sub WorksheetRemove( _
       ByVal Val As Variant, _
       ByVal ModuleName As String)
    ' Used in TableRoutines, WorksheetsClass

    Const RoutineName As String = Module_Name & "WorksheetRemove"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    pAllShts.Remove Val, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WorksheetRemove


