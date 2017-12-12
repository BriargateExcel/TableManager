Attribute VB_Name = "TableRoutines"
Option Explicit

Private Const Module_Name As String = "TableRoutines."

Private pAllTbls As TableManager.TablesClass

Public Sub ExtendDataValidationThroughAllTables(ByVal Wkbk As Workbook)
    Dim CurrentSheet As Worksheet
    Set CurrentSheet = MainWorkbook.ActiveSheet
    
    Dim RowCount As Long
    RowCount = 0
    Dim Tbl As TableManager.TableClass
    Dim I As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        ExtendDataValidationDownTable Tbl
        MainWorkbook.Worksheets(Tbl.WorksheetName).Activate
        Tbl.FirstCell.Select
    Next I
    
    CurrentSheet.Activate

End Sub

Private Sub ExtendDataValidationDownTable(ByVal Tbl As TableManager.TableClass)
    
    Dim I As Long
    Dim CopyRange As Range
    
    For I = 1 To Tbl.CellCount
        Set CopyRange = Tbl.ColumnRange(I)
        CopyRange(1, 1).Copy
        CopyRange.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next I
    
    Application.CutCopyMode = False

End Sub

Public Sub BuildTable( _
       ByVal WS As TableManager.WorksheetClass, _
       ByVal TblObj As ListObject)
    
    Dim Tbl As Variant
    Dim Frm As TableManager.FormClass
    
    Const RoutineName As String = Module_Name & "BuildTable"
    On Error GoTo ErrorHandler
    
    ' Gather the table data
    Set Tbl = New TableManager.TableClass
    Tbl.Name = TblObj.Name
    Set Tbl.Table = TblObj
    If Tbl.CollectTableData(WS, Tbl) Then
        Set Frm = New TableManager.FormClass
        TableManager.TableAdd Tbl, Module_Name
        
        Set Frm.FormObj = Frm.BuildForm(Tbl)
        Set Tbl.Form = Frm
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildTable

Public Function TableDataCollected() As Boolean
    On Error Resume Next
    TableDataCollected = (pAllTbls.Count <> 0)
    TableDataCollected = (Err.Number = 0)
End Function

Public Sub BuildParameterTableOnWorksheet(ByVal Wkbk As Workbook)
    Dim Tary As Variant
    
    Tary = BuildTableDataDescriptionArray
    
    BuildParameterWorksheet Wkbk, Tary, UBound(Tary, 1)
    
End Sub

Private Function BuildTableDataDescriptionArray() As Variant
    
    Dim TitleArray As Variant
    TitleArray = Array("Table Name", "Cell Name", "Cell Header Text", _
                       "Cell Type", "Operator", "Alert Style", "Formula 1", _
                       "Formula 2", "Ignore Blanks", "Show Input Message", _
                       "Input Title", "Input Message", "Show Error Message", _
                       "Error Title", "Error Message")
    
    Dim RowCount As Long
    RowCount = 0
    Dim Tbl As TableManager.TableClass
    Dim I As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        RowCount = RowCount + Tbl.CellCount
    Next I
    
    Dim Tary As Variant
    ReDim Tary(RowCount, UBound(TitleArray))
    For I = 0 To UBound(TitleArray)
        Tary(0, I) = TitleArray(I)
    Next I
    
    RowCount = 1
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        BuildRow Tary, Tbl, RowCount
    Next I
    
    BuildTableDataDescriptionArray = Tary
 
End Function

Private Sub BuildRow( _
        ByRef Tary As Variant, _
        ByVal Tbl As TableManager.TableClass, _
        ByRef RowCount As Long)
    
    Dim CellTypes As Variant
    CellTypes = Array("xlValidateInputOnly", "xlValidateWholeNumber", _
                      "xlValidateDecimal", "xlValidateList", "xlValidateDate", _
                      "xlValidateTime", "xlValidateTextLength", "xlValidateCustom")
    
    Dim Operators As Variant
    Operators = Array(vbNullString, "xlBetween", "xlNotBetween", "xlEqual", _
                      "xlNotEqual", "xlGreater", "xlLess", "xlGreaterEqual", _
                      "xlLessEqual")
    
    Dim AlertStyle As Variant
    AlertStyle = Array(vbNullString, "xlValidAlertStop", "xlValidAlertWarning", _
                       "xlValidAlertInformation")
    
    Dim Cll As TableManager.CellClass
    Dim J As Long
    For J = 0 To Tbl.CellCount - 1
        Set Cll = Tbl.TableCells.Item(J)
        Tary(RowCount, 0) = Tbl.Name
        Tary(RowCount, 1) = Cll.Name
        Tary(RowCount, 2) = Cll.HeaderText
        Tary(RowCount, 3) = CellTypes(Cll.CellType)
        PopulateValidationData Tary, RowCount, Operators, Cll, AlertStyle
        Tary(RowCount, 8) = IIf(Cll.IgnoreBlank, "True", "False")
        Tary(RowCount, 9) = Cll.ShowInput
        Tary(RowCount, 10) = Cll.InputTitle
        Tary(RowCount, 11) = Cll.InputMessage
        Tary(RowCount, 12) = Cll.ShowError
        Tary(RowCount, 13) = Cll.ErrorTitle
        Tary(RowCount, 14) = Cll.ErrorMessage
        RowCount = RowCount + 1
    Next J
End Sub
Private Sub PopulateValidationData( _
        ByRef Tary As Variant, _
        ByVal RowCount As Long, _
        ByVal Operators As Variant, _
        ByVal Cll As TableManager.CellClass, _
        ByVal AlertStyle As Variant)
    
    If Cll.CellType <> xlValidateInputOnly Then
        If Cll.CellType <> xlValidateList Then
            Tary(RowCount, 4) = Operators(Cll.Operator)
        End If
        Tary(RowCount, 5) = AlertStyle(Cll.ValidAlertStyle)
        If Left$(Cll.ValidationFormula1, 1) = "=" Then
            Tary(RowCount, 6) = "''" & Cll.ValidationFormula1
        Else
            Tary(RowCount, 6) = Cll.ValidationFormula1
        End If
        Tary(RowCount, 7) = Cll.ValidationFormula2
    End If
End Sub                                          ' PopulateValidationData

Private Sub BuildParameterWorksheet( _
        ByVal Wkbk As Workbook, _
        ByVal Tary As Variant, _
        ByVal RowCount As Long)
    
    Dim Rng As String
    
    With Wkbk
        If Not Contains(.Worksheets, "Parameters") Then
            .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).Name = "Parameters"
        End If
        With .Worksheets("Parameters")
            If Contains(.ListObjects, "ParameterTable") Then
                .ListObjects("ParameterTable").Delete
            End If
            .Range("$A$1").Resize(UBound(Tary, 1), UBound(Tary, 2)).Value = Tary
    
            Rng = "$A$1:" & ConvertToLetter(UBound(Tary, 2)) & RowCount - 1
            .ListObjects.Add(xlSrcRange, .Range(Rng), , xlYes).Name = "ParameterTable"
    
            .Activate
            .Range("A2").Select
        End With
    End With
    ActiveWindow.FreezePanes = True
    ActiveSheet.Range(Rng).EntireColumn.AutoFit
    
'TODO Add validation to the cells of the ParameterTable
'TODO Use the Parameter table to designate columns as keys then check them for uniqueness
'TODO Use a Parameter table to help in adding an "Add New Element" option to the dropdown menus

End Sub

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TablesClass.", "EventClass.", "TableClass.", "TableRoutines.")
End Function                                     ' ModuleList

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

End Sub                                          ' TurnOnCellDescriptions

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

End Sub                                          ' TurnOffCellDescriptions

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
        Case "lbl":                              ' Do nothing
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

End Sub                                          ' PopulateTable

Public Function Table( _
       ByVal TableName As String, _
       ByVal ModuleName As String _
       ) As TableManager.TableClass

    Const RoutineName As String = Module_Name & "Table"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)

    Set Table = pAllTbls.Item(TableName)

End Function                                     ' Table

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

End Sub                                          ' TableAdd

Public Function TableCount(ByVal ModuleName As String) As Long
    Const RoutineName As String = Module_Name & "TableCount"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    TableCount = pAllTbls.Count
End Function                                     ' TableCount

Public Function TableExists( _
       ByVal Tbl As Variant, _
       ByVal ModuleName As String _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TableExists"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    TableExists = pAllTbls.Exists(Tbl)
End Function                                     ' TableExists

Public Function TableItem( _
       ByVal Tbl As Variant, _
       ByVal ModuleName As String _
       ) As Variant

    Const RoutineName As String = Module_Name & "TableItem"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set TableItem = pAllTbls.Item(Tbl)
End Function                                     ' TableItem

Public Sub TableRemove( _
       ByVal Val As Variant, _
       ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "TableRemove"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    pAllTbls.Remove Val
End Sub                                          ' TableRemove

Public Sub TableSetNewClass(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNewClass"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = New TableManager.TablesClass
End Sub                                          ' TableSetNewClass

Public Sub TableSetNewDict(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNewDict"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = New Scripting.Dictionary
End Sub                                          ' TableSetNewDict

Public Sub TableSetNothing(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "TableSetNothing"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllTbls = Nothing
End Sub                                          ' TableSetNothing


