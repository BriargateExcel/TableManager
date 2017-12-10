Attribute VB_Name = "TableRoutines"
Option Explicit

Private Const Module_Name As String = "TableRoutines."

Private pAllTbls As TableManager.TablesClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TablesClass.", "EventClass.", "TableClass.", "SharedRoutines.")
End Function ' ModuleList

Public Sub TurnOnCellDescriptions( _
    ByVal Tbl As TableManager.TableClass, _
    ByVal ModuleName As String)
    
    Dim Field As TableManager.CellClass
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim DBRange As Range: Set DBRange = Tbl.DBRange:
    Dim I As Long
    
    Const RoutineName As String = Module_Name & "TurnOnCellDescriptions"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)

    On Error GoTo ErrorHandler

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I)
        Field.ShowInput = True
        DBCol = Tbl.DBCol(Field.HeaderText)
        On Error Resume Next
        ' If a cell has no validation, it will raise a 1004 error
        ' Therefore, there is no Validation object to set
        DBRange(DBRow, DBCol).Validation.ShowInput = True
        On Error GoTo ErrorHandler
    Next I
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' TurnOnCellDescriptions

Public Sub TurnOffCellDescriptions( _
    ByVal Tbl As TableManager.TableClass, _
    ByVal ModuleName As String)
    
    Dim Field As TableManager.CellClass
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim I As Long
    
    Const RoutineName As String = Module_Name & "TurnOffCellDescriptions"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)

    On Error GoTo ErrorHandler

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I)
        Field.ShowInput = False
        DBCol = Tbl.DBCol(Field.HeaderText)
        On Error Resume Next
        ' If a cell has no validation, it will raise a 1004 error
        ' Therefore, there is no Validation object to set
        DBRange(DBRow, DBCol).Validation.ShowInput = False
        On Error GoTo ErrorHandler
    Next I
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' TurnOffCellDescriptions

Public Sub PopulateTable( _
    ByVal Tbl As TableManager.TableClass, _
    ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "PopulateTable"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)

    On Error GoTo ErrorHandler

    Dim Field As TableManager.CellClass
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim I As Long

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I)
        DBCol = Tbl.DBCol(Field.HeaderText)

        Field.ControlValue = DBRange(DBRow, DBCol)
        
        Select Case Left$(Field.FormControl.Name, 3)
        Case "lbl": ' Do nothing
        Case "val": DBRange(DBRow, DBCol) = Field.FormControl.Caption
        Case "fld": DBRange(DBRow, DBCol) = Field.FormControl.Text
        Case "cmb": DBRange(DBRow, DBCol) = Field.FormControl.Text
        Case "whl": DBRange(DBRow, DBCol) = Field.FormControl.Text
        Case "dat": DBRange(DBRow, DBCol) = Field.FormControl.Text
        Case Else
            MsgBox _
                "This is an illegal field type: " & Left$(Field.FormControl.Name, 3), _
                vbOKOnly Or vbExclamation, "Illegal Field Type"

        End Select
        
    Next I
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' PopulateTable

Public Function Table( _
    ByVal TableName As String, _
    ByVal ModuleName As String _
    ) As TableManager.TableClass

    Const RoutineName As String = Module_Name & "Table"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)

    Set Table = pAllTbls.Item(TableName)

End Function ' Table

Public Sub TableAdd( _
    ByVal Tbl As Variant, _
    ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "TableAdd"
    On Error GoTo ErrorHandler
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    pAllTbls.Add Tbl
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' TableAdd

Public Function TableCount(ByVal ModuleName As String) As Long
    Const RoutineName As String = Module_Name & "TableCount"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    TableCount = pAllTbls.Count
End Function ' TableCount

Public Function TableExists( _
    ByVal Tbl As Variant, _
    ByVal ModuleName As String _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TableExists"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    TableExists = pAllTbls.Exists(Tbl)
End Function ' TableExists

Public Function TableItem( _
    ByVal Tbl As Variant, _
    ByVal ModuleName As String _
    ) As Variant

    Const RoutineName As String = Module_Name & "TableItem"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set TableItem = pAllTbls.Item(Tbl)
End Function ' TableItem

Public Sub TableRemove( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "TableRemove"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    pAllTbls.Remove Val
End Sub ' TableRemove

Public Sub TableSetNewClass(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNewClass"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = New TableManager.TablesClass
End Sub ' TableSetNewClass

Public Sub TableSetNewDict(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNewDict"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = New Scripting.Dictionary
End Sub ' TableSetNewDict

Public Sub TableSetNothing(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNothing"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = Nothing
End Sub ' TableSetNothing





