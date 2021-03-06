Attribute VB_Name = "ParameterRoutines"
'@Folder("TableManager.Colors")

Option Explicit

Private Const Module_Name As String = "ParameterRoutines."

Public Function DarkestColorValue() As Long
    Const RoutineName As String = Module_Name & "DarkestColorValue"
    On Error GoTo ErrorHandler
    
    DarkestColorValue = FieldValue(GetMainWorkbook, _
                                   "ColorTable", _
                                   "Color Name", _
                                   "Darkest Color", _
                                   "Decimal Color Value", _
                                   &H8000000F)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Public Function LightestColorValue() As Long
    LightestColorValue = FieldValue(GetMainWorkbook, _
                                    "ColorTable", _
                                    "Color Name", _
                                    "Lightest Color", _
                                    "Decimal Color Value", _
                                    &H80000012)
End Function

Public Function FieldValue( _
       ByVal Wkbk As Workbook, _
       ByVal TableName As String, _
       ByVal SearchFieldName As String, _
       ByVal SearchFieldValue As String, _
       ByVal TargetFieldName As String, _
       ByVal DefaultValue As Variant _
       ) As Variant
    
    Const RoutineName As String = Module_Name & "FieldValue"
    On Error GoTo ErrorHandler
    
    Dim LocalFieldValue As Variant
    
    If FieldExistsInXLAM(TableName, SearchFieldName) Then
        LocalFieldValue = GetCellValue(TableName, SearchFieldName, SearchFieldValue, TargetFieldName)
    Else
        If FieldExistsOnWorksheet(Wkbk, TableName, SearchFieldName) Then
            Dim Tbl As ListObject
            Set Tbl = Wkbk.Worksheets("Parameters").ListObjects(TableName)
            
            LocalFieldValue = SearchTable(Tbl, SearchFieldName, SearchFieldValue, TargetFieldName)
            If LocalFieldValue = 0 Then FieldValue = DefaultValue
        Else
            LocalFieldValue = DefaultValue
        End If
    End If
    
    FieldValue = LocalFieldValue
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Private Function FieldExistsInXLAM( _
        ByVal TableName As String, _
        ByVal FieldName As String _
        ) As Boolean
 
    Const RoutineName As String = Module_Name & "FieldExistsInXLAM"
    On Error GoTo ErrorHandler
    
    FieldExistsInXLAM = False
    
    If TableExists(TableName, Module_Name) Then
        Dim Tbl As TableClass
        Set Tbl = Table(TableName, Module_Name)
        
        FieldExistsInXLAM = Tbl.TableCells.Exists(FieldName, Module_Name)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Private Function FieldExistsOnWorksheet( _
        ByVal Wkbk As Workbook, _
        ByVal TableName As String, _
        ByVal FieldName As String _
        ) As Boolean
    
    Const RoutineName As String = Module_Name & "FieldExistsOnWorksheet"
    On Error GoTo ErrorHandler
    
    If TableExistsOnWorksheet(TableName) Then
        Dim Tbl As ListObject
        Set Tbl = Wkbk.Worksheets("Parameters").ListObjects(TableName)
        
        On Error Resume Next
        FieldExistsOnWorksheet = (Application.WorksheetFunction.Match(FieldName, Tbl.HeaderRowRange, 0) <> 0)
        On Error GoTo 0
        
        FieldExistsOnWorksheet = (Err.Number = 0)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Private Function TableExistsOnWorksheet( _
        ByVal TableName As String _
        ) As Boolean
    
    Const RoutineName As String = Module_Name & "TableExistsOnWorksheet"
    On Error GoTo ErrorHandler
    
    TableExistsOnWorksheet = False
    If ParameterSheetExists Then
        On Error Resume Next
        TableExistsOnWorksheet = Contains(GetMainWorkbook.Worksheets("Parameters").ListObjects, TableName)
        If Err.Number <> 0 Then TableExistsOnWorksheet = False
        On Error GoTo 0
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Private Function ParameterSheetExists() As Boolean
    Const RoutineName As String = Module_Name & "ParameterSheetExists"
    On Error GoTo ErrorHandler
    
    ParameterSheetExists = Contains(GetMainWorkbook.Worksheets, "Parameters")
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
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


