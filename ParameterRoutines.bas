Attribute VB_Name = "ParameterRoutines"
Option Explicit

Private Const Module_Name As String = "ParameterRoutines."

Public Function DarkestColorValue() As Long
    DarkestColorValue = FieldValue("ColorTable", _
                                   "Color Name", _
                                   "Darkest Color", _
                                   "Decimal Color Value", _
                                   &H80000012)
End Function

Public Function LightestColorValue() As Long
    LightestColorValue = FieldValue("ColorTable", _
                                    "Color Name", _
                                    "Lightest Color", _
                                    "Decimal Color Value", _
                                    &H8000000F)
End Function

Public Function FieldValue( _
       ByVal TableName As String, _
       ByVal SearchFieldName As String, _
       ByVal SearchFieldValue As String, _
       ByVal TargetFieldName As String, _
       ByVal DefaultValue As Variant _
       ) As Variant
    
    Dim LocalFieldValue As Variant
    
    If FieldExistsInXLAM(TableName, SearchFieldName) Then
        LocalFieldValue = TableManager.GetCellValue(TableName, SearchFieldName, SearchFieldValue, TargetFieldName)
    Else
        If FieldExistsOnWorksheet(TableName, SearchFieldName) Then
            Dim Tbl As ListObject
            Set Tbl = GetMainWorkbook.Worksheets("Parameters").ListObjects(TableName)
            
            LocalFieldValue = SearchTable(Tbl, SearchFieldName, SearchFieldValue, TargetFieldName)
            If LocalFieldValue = 0 Then FieldValue = DefaultValue
        Else
            LocalFieldValue = DefaultValue
        End If
    End If
    
    FieldValue = LocalFieldValue
End Function

Private Function FieldExistsInXLAM( _
        ByVal TableName As String, _
        ByVal FieldName As String _
        ) As Boolean
 
    FieldExistsInXLAM = False
    
    If TableManager.TableExists(TableName, Module_Name) Then
        Dim Tbl As TableManager.TableClass
        Set Tbl = TableManager.Table(TableName, Module_Name)
        
        FieldExistsInXLAM = Tbl.TableCells.Exists(FieldName, Module_Name)
    End If
End Function

Private Function FieldExistsOnWorksheet( _
        ByVal TableName As String, _
        ByVal FieldName As String _
        ) As Boolean
    
    If TableExistsOnWorksheet(TableName) Then
        Dim Tbl As ListObject
        Set Tbl = GetMainWorkbook.Worksheets("Parameters").ListObjects(TableName)
        
        On Error Resume Next
        FieldExistsOnWorksheet = (Application.WorksheetFunction.Match(FieldName, Tbl.HeaderRowRange, 0) <> 0)
        FieldExistsOnWorksheet = (Err.Number = 0)
    End If
End Function

Private Function TableExistsOnWorksheet(ByVal TableName As String) As Boolean
    TableExistsOnWorksheet = False
    If ParameterSheetExists Then
        TableExistsOnWorksheet = Contains(GetMainWorkbook.Worksheets("Parameters").ListObjects, TableName)
    End If
End Function

Private Function ParameterSheetExists() As Boolean
    ParameterSheetExists = Contains(GetMainWorkbook.Worksheets, "Parameters")
End Function

Private Function SearchTable( _
        ByVal Tbl As ListObject, _
        ByVal KeyColumnName As String, _
        ByVal KeyValue As String, _
        ByVal DataColumnName As String _
        ) As Long
    
    Dim KeyColumn As Long
    On Error Resume Next
    KeyColumn = Application.WorksheetFunction.Match(KeyColumnName, Tbl.HeaderRowRange, 0)
    On Error Resume Next
    If Err.Number <> 0 Or KeyColumn = 0 Then
        SearchTable = 0
        Exit Function
    End If
    On Error GoTo 0
        
    Dim KeyRange As Range
    Set KeyRange = Tbl.ListColumns(KeyColumn).DataBodyRange
        
    Dim KeyRow As Long
    On Error Resume Next
    KeyRow = Application.WorksheetFunction.Match(KeyValue, KeyRange, 0)
    If Err.Number <> 0 Or KeyRow = 0 Then
        SearchTable = 0
        Exit Function
    End If
    On Error GoTo 0
        
    On Error Resume Next
    Dim DataColumn As Long
    DataColumn = Application.WorksheetFunction.Match(DataColumnName, Tbl.HeaderRowRange, 0)
    If Err.Number <> 0 Or DataColumn = 0 Then
        SearchTable = 0
        Exit Function
    End If
    On Error GoTo 0
            
    SearchTable = Tbl.DataBodyRange(KeyRow, DataColumn)
End Function


