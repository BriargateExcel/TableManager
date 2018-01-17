Attribute VB_Name = "WorksheetRoutines"
Option Explicit

Private Const Module_Name As String = "WorksheetRoutines."

Private pAllShts As TableManager.WorksheetsClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TableRoutines.")
End Function                                     ' ModuleList

Public Sub WorksheetAdd( _
       ByVal WS As Variant, _
       ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "WorksheetAdd"

    Debug.Assert InScope(ModuleList, Modulename)
    pAllShts.Add WS
End Sub                                          ' WorksheetAdd

Public Sub WorksheetSetNewClass(ByVal Modulename As String)
    Const RoutineName As String = Module_Name & "WorksheetSetNewClass"
    Debug.Assert InScope(ModuleList, Modulename)
    Set pAllShts = New TableManager.WorksheetsClass
End Sub                                          ' WorksheetSetNewClass

Public Function WkSht( _
       ByVal WorksheetName As String, _
       ByVal Modulename As String _
       ) As TableManager.WorksheetClass

    Const RoutineName As String = Module_Name & "WkSht"
    On Error GoTo ErrorHandler
    Debug.Assert InScope(ModuleList, Modulename)

    Set WkSht = pAllShts.Item(WorksheetName)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' Table


