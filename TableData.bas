Attribute VB_Name = "TableData"
Option Explicit

Private pAllTbls As TablesClass
Private pAllShts As WorksheetsClass

Public Function Table( _
    ByVal TableName As String _
    ) As TableClass
    
    Dim Tbl As TableClass
    
    On Error Resume Next
    Set Tbl = pAllTbls(TableName)
    If Err.Number <> 0 Then
        Auto_Open
        Set Tbl = pAllTbls(TableName)
    End If
    On Error GoTo 0
    
    Set Table = Tbl
    
End Function

Public Function TableColumn( _
    ByVal TableName As Variant, _
    ByVal ColumnName As String _
    ) As Range
    
    Dim Tbl As TableClass
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
        Set TableColumn = Tbl.Worksheet.ListObjects(Tbl.Name).ListColumns(ColumnName).DataBodyRange
    
End Function

Public Function TableColumnNumber( _
    ByVal TableName As Variant, _
    ByVal ColumnName As String _
    ) As Long
    
    Dim Tbl As TableClass
    Dim HdrRng As Range
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
    Set HdrRng = Tbl.Worksheet.ListObjects(Tbl.Name).HeaderRowRange
    
    TableColumnNumber = VBAMatch(ColumnName, HdrRng)
    
End Function

Public Function TableRow( _
    ByVal TableName As Variant, _
    ByVal KeyColumn As Variant, _
    ByVal KeyVal As Variant _
    ) As Range
    
    Dim ColRng As Range
    Dim RowNum As Long
    Dim Tbl As TableClass
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
    If VarType(KeyColumn) = vbString Then
        Set ColRng = TableColumn(Tbl, KeyColumn)
    Else
'       KeyColumn is of type Range
        Set ColRng = KeyColumn
    End If
    
    RowNum = VBAMatch(KeyVal, ColRng)
    Set TableRow = Tbl.Worksheet.ListObjects(Tbl.Name).ListRows(RowNum).Range
    
End Function

Public Function TableRowNumber( _
    ByVal TableName As Variant, _
    ByVal KeyColumn As Variant, _
    ByVal KeyVal As Variant _
    ) As Long
    
    Dim ColRng As Range
    Dim Tbl As TableClass
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
    If VarType(KeyColumn) = vbString Then
        Set ColRng = TableColumn(Tbl, KeyColumn)
    Else
'       KeyColumn is of type Range
        Set ColRng = KeyColumn
    End If
    
    TableRowNumber = VBAMatch(KeyVal, ColRng)
    
End Function

Public Function GetTableCell( _
    ByVal TableName As Variant, _
    ByVal KeyColumn As Variant, _
    ByVal KeyVal As Variant, _
    ByVal TargetColumn As String _
    ) As Variant
    
    Dim Tbl As TableClass
    Dim ColRng As Range
    Dim RowNum As Long
    Dim ColNum As Long
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
    If VarType(KeyColumn) = vbString Then
        Set ColRng = TableColumn(Tbl, KeyColumn)
    Else
'       KeyColumn is of type Range
        Set ColRng = KeyColumn
    End If
    
    RowNum = VBAMatch(KeyVal, ColRng)
    ColNum = TableColumnNumber(Tbl, TargetColumn)
    
    GetTableCell = Tbl.Worksheet.ListObjects(Tbl.Name).DataBodyRange(RowNum, ColNum)
    
End Function

Public Sub LetTableCell( _
    ByVal TableName As Variant, _
    ByVal KeyColumn As Variant, _
    ByVal KeyVal As Variant, _
    ByVal TargetColumn As String, _
    ByVal NewVal As Variant)
    
    Dim Tbl As TableClass
    Dim ColRng As Range
    Dim RowNum As Long
    Dim ColNum As Long
    
    If VarType(TableName) = vbString Then
        Set Tbl = Table(TableName)
    Else
'       TableName is of type TableClass
        Set Tbl = TableName
    End If
    
    If VarType(KeyColumn) = vbString Then
        Set ColRng = TableColumn(Tbl, KeyColumn)
    Else
'       KeyColumn is of type Range
        Set ColRng = KeyColumn
    End If
    
    RowNum = VBAMatch(KeyVal, ColRng)
    ColNum = TableColumnNumber(Tbl, TargetColumn)
    
    Tbl.Worksheet.ListObjects(Tbl.Name).DataBodyRange(RowNum, ColNum) = NewVal
    
End Sub

Public Function InScope(ByVal ModuleList As Variant) As Boolean
' Retrieves the name of the module where InScope is called
' Filters the name against the list of valid module names
' Returns true if the results of the Filter has more than -1 entries
    InScope = _
        (UBound( _
            Filter(ModuleList, _
                Application.VBE.ActiveCodePane.CodeModule.Name, _
                True, _
                CompareMethod.BinaryCompare) _
        ) > -1)

End Function

Private Function TableModuleList() As Variant
    TableModuleList = Array("TablesClass", "TableData")
End Function

Public Function TableExists(ByVal Val As Variant) As Boolean
    Debug.Assert InScope(TableModuleList)
    TableExists = pAllTbls.Exists(Val)
End Function

Public Function TableItem(ByVal Val As Variant) As Variant
    Debug.Assert InScope(TableModuleList)
    Set TableItem = pAllTbls(Val)
End Function

Public Sub TableSetNothing()
    Debug.Assert InScope(TableModuleList)
    Set pAllTbls = Nothing
End Sub

Public Sub TableSetNewDict()
    Debug.Assert InScope(TableModuleList)
    Set pAllTbls = New Scripting.Dictionary
End Sub

Public Sub TableSetNewClass()
    Debug.Assert InScope(TableModuleList)
    Set pAllTbls = New TablesClass
End Sub

Public Sub TableRemove(ByVal Val As Variant)
    Debug.Assert InScope(TableModuleList)
    pAllTbls.Remove Val
End Sub

Public Function TableCount() As Long
    Debug.Assert InScope(TableModuleList)
    TableCount = pAllTbls.Count
End Function

Public Sub TableAdd(ByVal Val As Variant)
    Debug.Assert InScope(TableModuleList)
    pAllTbls.Add Val
End Sub

Private Function WorksheetModuleList() As Variant
    WorksheetModuleList = Array("WorksheetsClass", "TableData")
End Function

Public Function WorksheetExists(ByVal Val As Variant) As Boolean
    Debug.Assert InScope(WorksheetModuleList)
    WorksheetExists = pAllShts.Exists(Val)
End Function

Public Function WorksheetItem(ByVal Val As Variant) As Variant
    Debug.Assert InScope(WorksheetModuleList)
    Set WorksheetItem = pAllShts(Val)
End Function

Public Sub WorksheetSetNothing()
    Debug.Assert InScope(WorksheetModuleList)
    Set pAllShts = Nothing
End Sub

Public Sub WorksheetSetNewDict()
    Debug.Assert InScope(WorksheetModuleList)
    Set pAllShts = New Scripting.Dictionary
End Sub

Public Sub WorksheetSetNewClass()
    Debug.Assert InScope(WorksheetModuleList)
    Set pAllShts = New WorksheetsClass
End Sub

Public Sub WorksheetRemove(ByVal Val As Variant)
    Debug.Assert InScope(WorksheetModuleList)
    pAllShts.Remove Val
End Sub

Public Function WorksheetCount() As Long
    Debug.Assert InScope(WorksheetModuleList)
    WorksheetCount = pAllShts.Count
End Function

Public Sub WorksheetAdd(ByVal Val As Variant)
    Debug.Assert InScope(WorksheetModuleList)
    pAllShts.Add Val
End Sub


