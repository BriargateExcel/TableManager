Attribute VB_Name = "WorksheetRoutines"
Option Explicit

Private Const Module_Name = "WorksheetRoutines."

Private pAllShts As TableManager.WorksheetsClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.")
End Function ' ModuleList

Public Sub WorksheetAdd( _
    ByVal WS As Variant, _
    ByVal ModuleName As String)
    
    Const RoutineName = Module_Name & "WorksheetAdd"

    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    pAllShts.Add WS
End Sub ' WorksheetAdd

Public Function WorksheetCount(ByVal ModuleName As String) As Long
    Const RoutineName = Module_Name & "WorksheetCount"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    WorksheetCount = pAllShts.Count
End Function ' WorksheetCount

Public Function WorksheetExists( _
    ByVal WS As Variant, _
    ByVal ModuleName As String _
    ) As Boolean

    Const RoutineName = Module_Name & "WorksheetExists"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    WorksheetExists = pAllShts.Exists(WS)
End Function ' WorksheetExists

Public Sub WorksheetSetNewClass(ByVal ModuleName As String)
    Const RoutineName = Module_Name & "WorksheetSetNewClass"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllShts = New TableManager.WorksheetsClass
End Sub ' WorksheetSetNewClass

