Attribute VB_Name = "WorksheetRoutines"
Option Explicit

Private pAllShts As WorksheetsClass

Public Sub WorksheetAdd( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)
    
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    pAllShts.Add Val
End Sub

Public Function WorksheetCount(ByVal ModuleName As String) As Long
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    WorksheetCount = pAllShts.Count
End Function

Public Function WorksheetExists( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As Boolean
    
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    WorksheetExists = pAllShts.Exists(Val)
End Function

Public Function WorksheetItem( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As Variant
    
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    Set WorksheetItem = pAllShts(Val)
End Function

Private Function WorksheetModuleList() As Variant
    WorksheetModuleList = Array("WorksheetsClass.", "Module1.")
End Function

Public Sub WorksheetRemove( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)
    
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    pAllShts.Remove Val
End Sub

Public Sub WorksheetSetNewClass(ByVal ModuleName As String)
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    Set pAllShts = New WorksheetsClass
End Sub

Public Sub WorksheetSetNewDict(ByVal ModuleName As String)
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    Set pAllShts = New Scripting.Dictionary
End Sub

Public Sub WorksheetSetNothing(ByVal ModuleName As String)
    Debug.Assert InScope(WorksheetModuleList, ModuleName)
    Set pAllShts = Nothing
End Sub

