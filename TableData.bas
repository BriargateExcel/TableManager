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
    
    TableRowNumber = VBAMatch(KeyVal, ColRng)
    
End Function

Public Function TableCell( _
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
    
    TableCell = Tbl.Worksheet.ListObjects(Tbl.Name).DataBodyRange(RowNum, ColNum)
    
End Function

Public Function TableInScope() As Boolean
' Retrieves the name of the module where InScope is called
' Filters the name against the list of valid module names
' Returns true if the results of the Filter has more than -1 entries
    TableInScope = _
        (UBound( _
            Filter(TableModuleList, _
                Application.VBE.ActiveCodePane.CodeModule.Name, _
                True, _
                CompareMethod.BinaryCompare) _
        ) > -1)

End Function

Private Function TableModuleList() As Variant
    TableModuleList = Array("TablesClass", "TableData")
End Function

Public Function TableExists(ByVal Val As Variant) As Boolean
    Debug.Assert TableInScope
    TableExists = pAllTbls.Exists(Val)
End Function

Public Function TableItem(ByVal Val As Variant) As Variant
    Debug.Assert TableInScope
    Set TableItem = pAllTbls(Val)
End Function

Public Sub TableSetNothing()
    Debug.Assert TableInScope
    Set pAllTbls = Nothing
End Sub

Public Sub TableSetNewDict()
    Debug.Assert TableInScope
    Set pAllTbls = New Scripting.Dictionary
End Sub

Public Sub TableSetNewClass()
    Debug.Assert TableInScope
    Set pAllTbls = New TablesClass
End Sub

Public Sub TableRemove(ByVal Val As Variant)
    Debug.Assert TableInScope
    pAllTbls.Remove Val
End Sub

Public Function TableCount() As Long
    Debug.Assert TableInScope
    TableCount = pAllTbls.Count
End Function

Public Sub TableAdd(ByVal Val As Variant)
    Debug.Assert TableInScope
    pAllTbls.Add Val
End Sub

Public Function WorksheetInScope() As Boolean
' Retrieves the name of the module where InScope is called
' Filters the name against the list of valid module names
' Returns true if the results of the Filter has more than -1 entries
    WorksheetInScope = _
        (UBound( _
            Filter(WorksheetModuleList, _
                Application.VBE.ActiveCodePane.CodeModule.Name, _
                True, _
                CompareMethod.BinaryCompare) _
        ) > -1)

End Function

Private Function WorksheetModuleList() As Variant
    WorksheetModuleList = Array("WorksheetsClass", "TableData")
End Function

Public Function WorksheetExists(ByVal Val As Variant) As Boolean
    Debug.Assert WorksheetInScope
    WorksheetExists = pAllShts.Exists(Val)
End Function

Public Function WorksheetItem(ByVal Val As Variant) As Variant
    Debug.Assert WorksheetInScope
    Set WorksheetItem = pAllShts(Val)
End Function

Public Sub WorksheetSetNothing()
    Debug.Assert WorksheetInScope
    Set pAllShts = Nothing
End Sub

Public Sub WorksheetSetNewDict()
    Debug.Assert WorksheetInScope
    Set pAllShts = New Scripting.Dictionary
End Sub

Public Sub WorksheetSetNewClass()
    Debug.Assert WorksheetInScope
    Set pAllShts = New WorksheetsClass
End Sub

Public Sub WorksheetRemove(ByVal Val As Variant)
    Debug.Assert WorksheetInScope
    pAllShts.Remove Val
End Sub

Public Function WorksheetCount() As Long
    Debug.Assert WorksheetInScope
    WorksheetCount = pAllShts.Count
End Function

Public Sub WorksheetAdd(ByVal Val As Variant)
    Debug.Assert WorksheetInScope
    pAllShts.Add Val
End Sub


