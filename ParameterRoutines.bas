Attribute VB_Name = "ParameterRoutines"
Option Explicit

Private Const Module_Name As String = "ParameterRoutines."

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

Public Function DarkestColorValue() As Long
    If ColorNameFieldExistsInXLAM Then
        DarkestColorValue = TableManager.GetCellValue("ColorTable", "Color Name", "Darkest Color", "Decimal Color Value")
    Else
        If ColorTableExistsOnWorksheet Then
            Dim Tbl As ListObject
            Set Tbl = MainWorkbook.Worksheets("Parameters").ListObjects("ColorTable")
            
            DarkestColorValue = SearchTable(Tbl, "Color Name", "Darkest Color", "Decimal Color Value")
            If DarkestColorValue = 0 Then DarkestColorValue = &H80000007 ' Black
        Else
            DarkestColorValue = &H80000007       ' Black
        End If
    End If
End Function

Public Function LightestColorValue() As Long
    If ColorNameFieldExistsInXLAM Then
        LightestColorValue = TableManager.GetCellValue("ColorTable", "Color Name", "Lightest Color", "Decimal Color Value")
    Else
        If ColorTableExistsOnWorksheet Then
            Dim Tbl As ListObject
            Set Tbl = MainWorkbook.Worksheets("Parameters").ListObjects("ColorTable")
            
            LightestColorValue = SearchTable(Tbl, "Color Name", "Lightest Color", "Decimal Color Value")
            If LightestColorValue = 0 Then LightestColorValue = &H8000000F ' Light gray
        Else
            LightestColorValue = &H8000000F      ' Light gray
        End If
    End If
End Function

Private Function ColorNameFieldExistsInXLAM() As Boolean
    If ColorTableExistsOnWorksheet Then
        Dim Tbl As ListObject
        Set Tbl = MainWorkbook.Worksheets("Parameters").ListObjects("ColorTable")
        
        ColorNameFieldExistsInXLAM = Contains(Tbl.HeaderRowRange, "Color Name")
    Else
        ColorNameFieldExistsInXLAM = False
    End If
End Function

Private Function ColorTableExistsOnWorksheet() As Boolean
    If ParameterSheetExists Then
        ColorTableExistsOnWorksheet = Contains(MainWorkbook.Worksheets("Parameters").ListObjects, "ColorTable")
    Else
        ColorTableExistsOnWorksheet = False
    End If
End Function

Private Function ParameterSheetExists() As Boolean
    ParameterSheetExists = Contains(MainWorkbook.Worksheets, "Parameters")
End Function

Private Function ColorTableExistsInXLAM() As Boolean
    ColorTableExistsInXLAM = TableManager.TableExists("ColorTable", Module_Name)
End Function


