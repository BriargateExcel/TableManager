Attribute VB_Name = "TableRoutines"
'@Folder("TableManager.Tables")

Option Explicit

Private Const Module_Name As String = "TableRoutines."

Private pAllTbls As TablesClass

Private Const ListOfTypes As String = "xlValidateList,xlValidateInputOnly," & _
"xlValidateWholeNumber,xlValidateDecimal,xlValidateList,xlValidateDate," & _
"xlValidateTime,xlValidateTextLength,xlValidateCustom"

Private Const ListOfOperators As String = "xlBetween,xlNotBetween,xlEqual,xlNotEqual," & _
"xlGreater,xlLess,xlGreaterEqual,xlLessEqual"

Private Const ListOfAlertStyles As String = "xlValidAlertStop,xlValidAlertWarning," & _
"xlValidAlertInformation"

Private Const ListOfTruefFalse As String = "True,False"

Private Const ListOfYesNo As String = "Yes,No"

Public Function GetCellValue( _
       ByVal TableName As String, _
       ByVal KeyColumnName As String, _
       ByVal KeyValue As String, _
       ByVal DataColumnName As String _
       ) As Variant
' Used in ParameterRoutines

    Dim Tbl As TableClass
    Set Tbl = Table(TableName, Module_Name)
    
    Dim TableRow As Long
    TableRow = Tbl.DBRowNumber(KeyColumnName, KeyValue)
    If TableRow = 0 Then
        Err.Raise 1, "TableRoutines.GetCellValue", "Fatal error. KeyValue not found."
    End If
    
    Dim TableColumn As Long
    TableColumn = Tbl.DBColNumber(DataColumnName)
    If TableColumn = 0 Then
        Err.Raise 1, "TableRoutines.GetCellValue", "Fatal error. DataColumnName not found."
    End If
    
    GetCellValue = Tbl.DBRange(TableRow, TableColumn)
    
End Function

Private Function ParameterDescriptionArray() As Variant
    
    Dim PDA As Variant                           ' Parameter Description Array
    PDA = Array( _
          Array("Table Name", "xlValidateInputOnly"), _
                Array("Cell Header Text", "xlValidateInputOnly"), _
                      Array("Key", "xlValidateList", ListOfYesNo), _
                      Array("Cell Name", "xlValidateInputOnly"), _
                            Array("Cell Type", "xlValidateList", ListOfTypes, "WrapText"), _
                            Array("Operator", "xlValidateList", ListOfOperators, "WrapText"), _
                            Array("Alert Style", "xlValidateList", ListOfAlertStyles, "WrapText"), _
                            Array("Formula 1", "xlValidateInputOnly", , "WrapText"), _
                            Array("Formula 2", "xlValidateInputOnly", , "WrapText"), _
                            Array("Ignore Blanks", "xlValidateList", ListOfTruefFalse), _
                            Array("Show Input Message", "xlValidateList", ListOfTruefFalse), _
                            Array("Input Title", "xlValidateInputOnly"), _
                                  Array("Input Message", "xlValidateInputOnly", , "WrapText"), _
                                  Array("Show Error Message", "xlValidateList", ListOfTruefFalse), _
                                  Array("Error Title", "xlValidateInputOnly"), _
                                        Array("Error Message", "xlValidateInputOnly", , "WrapText") _
                                        )
    
    ParameterDescriptionArray = PDA
    
End Function                                     ' ParameterDescriptionArray

Private Sub BuildRow( _
        ByRef Tary As Variant, _
        ByVal Tbl As TableClass, _
        ByRef RowCount As Long)
    
    Dim CellTypes As Variant
    CellTypes = Split(ListOfTypes, ",")
    
    Dim Operators As Variant
    Operators = Split(ListOfOperators, ",")
    
    Dim AlertStyle As Variant
    AlertStyle = Split(ListOfAlertStyles, ",")
    
    Dim Cll As CellClass
    Dim J As Long
    For J = 0 To Tbl.CellCount - 1
        Set Cll = Tbl.TableCells.Item(J, Module_Name)
        Tary(RowCount, 0) = Tbl.Name
        Tary(RowCount, 1) = Cll.Name
        Tary(RowCount, 3) = Cll.HeaderText
        Tary(RowCount, 4) = CellTypes(Cll.CellType + 1)
        PopulateValidationData Tary, RowCount, Operators, Cll, AlertStyle
        Tary(RowCount, 9) = IIf(Cll.IgnoreBlank, "True", "False")
        Tary(RowCount, 10) = Cll.ShowInput
        Tary(RowCount, 11) = Cll.InputTitle
        Tary(RowCount, 12) = Cll.InputMessage
        Tary(RowCount, 13) = Cll.ShowError
        Tary(RowCount, 14) = Cll.ErrorTitle
        Tary(RowCount, 15) = Cll.ErrorMessage
        RowCount = RowCount + 1
    Next J
End Sub

Private Function BuildTableDataDescriptionArray() As Variant
    
    Dim TitleArray As Variant
    TitleArray = ParameterDescriptionArray
    
    ' Calculate the number of rows required to store all the cell data in all the tables
    Dim RowCount As Long
    RowCount = 0
    Dim Tbl As TableClass
    Dim I As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        RowCount = RowCount + Tbl.CellCount
    Next I
    
    ' Populate the first row with column headers
    Dim Tary As Variant
    ReDim Tary(RowCount, UBound(TitleArray))
    For I = 0 To UBound(TitleArray)
        Tary(0, I) = TitleArray(I)(0)
    Next I
    
    ' Populate the remaining rows with data; one row per cell
    RowCount = 1
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        BuildRow Tary, Tbl, RowCount
    Next I
    
    BuildTableDataDescriptionArray = Tary
 
End Function

Private Sub BuildValidateInput( _
        ByVal Tbl As ListObject, _
        ByVal PDA As Variant, _
        ByVal NumCol As Long)
    
    Dim TopCell As Range
    
    Set TopCell = Tbl.DataBodyRange(1, NumCol + 1)
    TopCell.Validation.Delete
    TopCell.Validation.Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop
    SetCommonValidationParameters Tbl, PDA, NumCol
End Sub

Private Sub BuildValidateList( _
        ByVal Tbl As ListObject, _
        ByVal PDA As Variant, _
        ByVal NumCol As Long)

    Dim TopCell As Range
    
    Set TopCell = Tbl.DataBodyRange(1, NumCol + 1)
    TopCell.Validation.Delete
    TopCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=PDA(NumCol)(2)
    SetCommonValidationParameters Tbl, PDA, NumCol
    
End Sub

Private Sub AddValidationToParameterTable(ByRef Tbl As ListObject)

    Dim I As Long
    Dim PDA As Variant
    PDA = ParameterDescriptionArray
    
    For I = 0 To UBound(PDA, 1)
        If PDA(I)(1) = "xlValidateInputOnly" Then
            BuildValidateInput Tbl, PDA, I
        ElseIf PDA(I)(1) = "xlValidateList" Then
            BuildValidateList Tbl, PDA, I
        Else
            MsgBox "The Parameter Description Array has a bad value", _
                   vbOKOnly Or vbCritical, _
                   "Error found in TableRoutines.AddValidationToParameterTable"
        End If
    
    Next I
    
    Tbl.DataBodyRange(1, 1).Select
    
End Sub

Private Sub ExtendDataValidationDownTable(ByVal Tbl As TableClass)
    
    Const RoutineName As String = Module_Name & "ExtendDataValidationDownTable"
    On Error GoTo ErrorHandler
    
    Dim I As Long
    Dim CopyRange As Range
    
    For I = 1 To Tbl.CellCount
        Set CopyRange = Tbl.ColumnRange(I)
        CopyRange(1, 1).Copy
        CopyRange.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next I
    
    Application.CutCopyMode = False

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

Public Sub BuildParameterTableOnWorksheet(ByVal Wkbk As Workbook)
' Used in Main

    ' Assumes that all tables start in Row 1
    
    Const RoutineName As String = Module_Name & "BuildParameterTableOnWorksheet"
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False

    Dim Tary As Variant
    Tary = BuildTableDataDescriptionArray
    
    Dim RowCount As Long
    RowCount = UBound(Tary, 1)
    
    Dim ColumnCount As Long
    ColumnCount = UBound(Tary, 2)
    
    With Wkbk
        ' Ensure there's a sheet called "Parameters"
        If Not Contains(.Worksheets, "Parameters") Then
            .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).Name = "Parameters"
        End If
        
        ' Determine where to put the ParameterTable
        Dim UpperLeftCorner As String
        Dim FreezePoint As String
        Dim ColumnNumber As Long
        Dim ColumnLetter As String
        
        If Contains(.Worksheets("Parameters").ListObjects, "ParameterTable") Then
            Dim Header As Range
            Set Header = .Worksheets("Parameters").ListObjects("ParameterTable").HeaderRowRange
            
            ColumnNumber = Header.Columns(1).Column
            ColumnLetter = ConvertToLetter(ColumnNumber)
            UpperLeftCorner = ColumnLetter & "1"
            FreezePoint = ColumnLetter & "2"
        Else
            ColumnNumber = FindLastColumnNumber(1, .Worksheets("Parameters")) + 2
            
            ColumnLetter = ConvertToLetter(ColumnNumber)
            UpperLeftCorner = ColumnLetter & "1"
            FreezePoint = ColumnLetter & "2"
        End If
        
        With .Worksheets("Parameters")
            Dim Rng As String
            If Contains(.ListObjects, "ParameterTable") Then
                .ListObjects("ParameterTable").Delete
                SetInitializing
                TableRemove "ParameterTable", Module_Name
                ReSetInitializing
            End If
            .Range(UpperLeftCorner).Resize(RowCount + 1, ColumnCount + 1).Value = Tary
    
            Rng = UpperLeftCorner & ":" & ConvertToLetter(ColumnNumber + ColumnCount) & RowCount
            .ListObjects.Add(xlSrcRange, .Range(Rng), , xlYes).Name = "ParameterTable"
    
            .Activate
            .Range(FreezePoint).Select
            
            AddValidationToParameterTable .ListObjects("ParameterTable")
            
        End With                                 ' .Worksheets("Parameters")
        BuildTable Wkbk, .ListObjects("ParameterTable"), Module_Name
    End With                                     ' Wkbk
        
     
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = False
        
    'TODO Use the Parameter table to designate columns as keys then check them for uniqueness
    'TODO Use a Parameter table to help in adding an "Add New Element" option to the dropdown menus
    ' This will be a challenge because the Parameter table is the last to be built
    ' and all the forms are already built

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

'Private Sub BuildTableAndForm()
'    Const RoutineName As String = Module_Name & "BuildTableAndForm"
'    On Error GoTo ErrorHandler
'
'
'    '@Ignore LineLabelNotUsed
'Done:
'    Exit Sub
'ErrorHandler:
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub

Private Sub SetCommonValidationParameters( _
        ByRef Tbl As ListObject, _
        ByVal PDA As Variant, _
        ByVal ColNum As Long)
    
    Dim Cll As Range
    Set Cll = Tbl.DataBodyRange(1, ColNum + 1)
    
    Cll.Locked = False
    Cll.FormulaHidden = False
    
    With Cll.Validation
        .InCellDropDown = (PDA(ColNum)(1) = "xlValidateList")
        .IgnoreBlank = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
    
    Dim CopyRange As Range
    
    Set CopyRange = Tbl.ListColumns(ColNum + 1).DataBodyRange
    Cll.Copy
    CopyRange.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    On Error Resume Next
    Dim Wrapper As String
    Wrapper = PDA(ColNum)(3)
    If Err.Number <> 0 Then Exit Sub
    If Wrapper = "WrapText" Then
        CopyRange.WrapText = True
    End If
    
End Sub

Public Sub ExtendDataValidationThroughAllTables(ByVal Wkbk As Workbook)
' Used in Main

    Const RoutineName As String = Module_Name & "ExtendDataValidationThroughAllTables"
    On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    
    Dim CurrentSheet As Worksheet
    Set CurrentSheet = Wkbk.ActiveSheet
    
    Dim CurrentCell As Range
    Set CurrentCell = ActiveCell
    
    Dim Tbl As TableClass
    Dim I As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        ExtendDataValidationDownTable Tbl
        Wkbk.Worksheets(Tbl.WorksheetName).Activate
        Tbl.FirstCell.Select
    Next I
    
    CurrentSheet.Activate
    CurrentCell.Select
    
    Application.ScreenUpdating = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

Public Sub BuildTable( _
        ByVal Wkbk As Workbook, _
       ByVal TblObj As ListObject, _
       ByVal ModuleName As String)
' Used in XLAM_Module, TableRoutines
       
    Const RoutineName As String = Module_Name & "BuildTable"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    SetInitializing
    
    ' Gather the table data
    Dim Tbl As Variant
    Set Tbl = New TableClass
    Tbl.Name = TblObj.Name
    Set Tbl.Table = TblObj
    
    Dim Sht As Worksheet
    Set Sht = Wkbk.Worksheets(TblObj.Parent.Name)
    
    If Tbl.CollectTableData(Wkbk, Tbl, Module_Name) Then
        Dim Frm As FormClass
        Set Frm = New FormClass
        TableAdd Tbl, Module_Name
            
        Set Frm.FormObj = Frm.BuildForm(Wkbk, Tbl, Module_Name)
        Set Tbl.Form = Frm
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' BuildTable

Public Function TableDataCollected() As Boolean
' Used in Main
    On Error Resume Next
    TableDataCollected = (pAllTbls.Count <> 0)
    TableDataCollected = (Err.Number = 0)
End Function

Private Sub PopulateValidationData( _
        ByRef Tary As Variant, _
        ByVal RowCount As Long, _
        ByVal Operators As Variant, _
        ByVal Cll As CellClass, _
        ByVal AlertStyle As Variant)
    
    If Cll.CellType <> xlValidateInputOnly Then
        If Cll.CellType <> xlValidateList Then
            Tary(RowCount, 5) = Operators(Cll.Operator)
        End If
        Tary(RowCount, 6) = AlertStyle(Cll.ValidAlertStyle + 1)
        If Left$(Cll.ValidationFormula1, 1) = "=" Then
            Tary(RowCount, 7) = "''" & Cll.ValidationFormula1
        Else
            Tary(RowCount, 7) = Cll.ValidationFormula1
        End If
        Tary(RowCount, 8) = Cll.ValidationFormula2
    End If
End Sub                                          ' PopulateValidationData

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "TablesClass.", "EventClass.", "TableClass.", "TableRoutines.", "ParameterRoutines.", "DataBaseRoutines.")
End Function                                     ' ModuleList

Public Sub TurnOnCellDescriptions( _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)
    
    Dim Field As CellClass
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim DBRange As Range: Set DBRange = Tbl.DBRange:
    Dim I As Long
    
    Const RoutineName As String = Module_Name & "TurnOnCellDescriptions"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)

    On Error GoTo ErrorHandler

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
        Field.ShowInput = True
        DBCol = Tbl.SelectedDBCol(Field.HeaderText)
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
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)
    
    Dim Field As CellClass
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim I As Long
    
    Const RoutineName As String = Module_Name & "TurnOffCellDescriptions"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)

    On Error GoTo ErrorHandler

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
        Field.ShowInput = False
        DBCol = Tbl.SelectedDBCol(Field.HeaderText)
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
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)
' Used in EventClass

    Const RoutineName As String = Module_Name & "PopulateTable"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)

    On Error GoTo ErrorHandler

    Dim Field As CellClass
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim I As Long

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
        DBCol = Tbl.SelectedDBCol(Field.HeaderText)
        If DBCol = 0 Then
            Err.Raise 1, "TableRoutines.PopulateTable", "Fatal error. HeaderText not found."
        End If

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
       ) As TableClass
' Used in TableRoutines, ParameterRoutines, EventClass

    Const RoutineName As String = Module_Name & "Table"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)

    Set Table = pAllTbls.Item(TableName, Module_Name)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' Table

Private Sub TableAdd( _
       ByVal Tbl As Variant, _
       ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "TableAdd"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    pAllTbls.Add Tbl, Module_Name
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' TableAdd

Public Function TableCount(ByVal ModuleName As String) As Long
    Const RoutineName As String = Module_Name & "TableCount"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    TableCount = pAllTbls.Count

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' TableCount

Public Function TableExists( _
       ByVal Tbl As Variant, _
       ByVal ModuleName As String _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TableExists"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    TableExists = pAllTbls.Exists(Tbl, Module_Name)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' TableExists

Public Function TableItem( _
       ByVal Tbl As Variant, _
       ByVal ModuleName As String _
       ) As Variant

    Const RoutineName As String = Module_Name & "TableItem"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set TableItem = pAllTbls.Item(Tbl, Module_Name)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' TableItem

Public Sub TableRemove( _
       ByVal Val As Variant, _
       ByVal ModuleName As String)
' Used in TableRoutines, WorksheetsClass

    Const RoutineName As String = Module_Name & "TableRemove"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    pAllTbls.Remove Val, Module_Name

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' TableRemove

Public Sub TableSetNewClass(ByVal ModuleName As String)
' Used in XLAM_Module

    Const RoutineName As String = Module_Name & "TableSetNewClass"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllTbls = New TablesClass

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' TableSetNewClass

Public Sub TableSetNewDict(ByVal ModuleName As String)
' Used in WorksheetsClass

    Const RoutineName As String = Module_Name & "TableSetNewDict"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllTbls = New Scripting.Dictionary

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' TableSetNewDict

Public Sub TableSetNothing(ByVal ModuleName As String)
' Used in WorksheetsClass

    Const RoutineName As String = Module_Name & "TableSetNothing"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    Set pAllTbls = Nothing

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' TableSetNothing

Public Sub CopyToTable( _
    ByVal Wkbk As Workbook, _
       ByVal TableName As String, _
       ByVal Ary As Variant)
' Used in CSVRoutines

    Const RoutineName As String = Module_Name & "CopyToTable"
    On Error GoTo ErrorHandler
    
    ' Copy the file to the table
    Dim Sht As Worksheet
    Set Sht = Wkbk.Worksheets(Table(TableName, Module_Name).WorksheetName)
    
    Dim UpperLeftRange As Range
    Dim Tbl As TableClass
    Set Tbl = Table(TableName, Module_Name)
    Dim UpperLeftRng As Range
    Set UpperLeftRng = Tbl.FirstCell
    
    Set UpperLeftRange = UpperLeftRng.Offset(-1, 0)
    
    UpperLeftRange.Resize(UBound(Ary, 1), UBound(Ary, 2)) = Ary
    
    ' Re-establish the table
    Dim UpperLeft As String
    Dim LowerRight As String
    UpperLeft = UpperLeftRng.Offset(-1, 0).Address
    LowerRight = ConvertToLetter(UBound(Ary, 2)) & UBound(Ary, 1)
    Sht.ListObjects.Add(xlSrcRange, Sht.Range(UpperLeft & ":" & LowerRight), , xlYes).Name = TableName
    Set Tbl.Table = Sht.ListObjects(TableName)
    
    ' Re-establish the lock and data validation
    Dim I As Long
    Dim Cll As CellClass
    Dim ColRng As Range
    For I = 0 To Tbl.CellCount - 1
        Set Cll = Tbl.TableCells.Item(I, Module_Name)
        Set ColRng = Tbl.DBColRange(Tbl, Cll.HeaderText)
        
        If Cll.Locked Then
            ColRng.Locked = True
        Else
            ColRng.Locked = False
        End If
        
        If Cll.CellType <> xlValidateInputOnly Then
            With ColRng.Validation
                .Delete
                .Add Cll.CellType, Cll.ValidAlertStyle, Cll.Operator, Cll.ValidationFormula1, Cll.ValidationFormula2
                
                .IgnoreBlank = Cll.IgnoreBlank
                .InCellDropDown = Cll.InCellDropDown
            
                .ShowInput = Cll.ShowInput
                .InputTitle = Cll.InputTitle
                .InputMessage = Cll.InputMessage
            
                .ShowError = Cll.ShowError
                .ErrorTitle = Cll.ErrorTitle
                .ErrorMessage = Cll.ErrorMessage
            
            End With                             ' ColRng.Validation
        End If
        
    Next I

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub


