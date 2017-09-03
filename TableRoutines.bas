Attribute VB_Name = "TableRoutines"
Option Explicit

Private pAllTbls As TablesClass

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

Public Function Table( _
    ByVal TableName As String _
    ) As TableClass
    
    Dim Tbl As TableClass
    
'    On Error Resume Next
    Set Tbl = pAllTbls(TableName)
'    If Err.Number <> 0 Then
'        Auto_Open
'        Set Tbl = pAllTbls(TableName)
'    End If
'    On Error GoTo 0
    
    Set Table = Tbl
    
End Function

Public Sub TableAdd( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)
    
    Debug.Assert InScope(TableModuleList, ModuleName)
    pAllTbls.Add Val
End Sub

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

Public Function TableCount(ByVal ModuleName As String) As Long
    Debug.Assert InScope(TableModuleList, ModuleName)
    TableCount = pAllTbls.Count
End Function

Public Function TableExists( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As Boolean
    
    Debug.Assert InScope(TableModuleList, ModuleName)
    TableExists = pAllTbls.Exists(Val)
End Function

Public Function TableItem( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As Variant
    
    Debug.Assert InScope(TableModuleList, ModuleName)
    Set TableItem = pAllTbls(Val)
End Function

Private Function TableModuleList() As Variant
    TableModuleList = Array("TablesClass.", "Module1.", "WorksheetClass.", "EventHandler.")
End Function

Public Sub TableRemove( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)
    
    Debug.Assert InScope(TableModuleList, ModuleName)
    pAllTbls.Remove Val
End Sub

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

Public Sub TableSetNewClass(ByVal ModuleName As String)
    Debug.Assert InScope(TableModuleList, ModuleName)
    Set pAllTbls = New TablesClass
End Sub

Public Sub TableSetNewDict(ByVal ModuleName As String)
    Debug.Assert InScope(TableModuleList, ModuleName)
    Set pAllTbls = New Scripting.Dictionary
End Sub

Public Sub TableSetNothing(ByVal ModuleName As String)
    Debug.Assert InScope(TableModuleList, ModuleName)
    Set pAllTbls = Nothing
End Sub

