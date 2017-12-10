Attribute VB_Name = "TableRoutines"
Option Explicit

Private Const Module_Name As String = "TableRoutines."

Private pAllTbls As TableManager.TablesClass

Public Sub BuildTableDataDescriptionArray(ByVal WkBkName As String)
    
    Dim TitleArray As Variant
    TitleArray = Array("Table Name", "Cell Name", "Cell Header Text", _
                       "Cell Type", "Operator", "Alert Style", "Formula 1", _
                       "Formula 2", "Ignore Blanks", "Show Input Message", _
                       "Input Title", "Input Message", "Show Error Message", _
                       "Error Title", "Error Message")
    
    Dim CellTypes As Variant
    CellTypes = Array("xlValidateInputOnly", "xlValidateWholeNumber", _
                      "xlValidateDecimal", "xlValidateList", "xlValidateDate", _
                      "xlValidateTime", "xlValidateTextLength", "xlValidateCustom")
    
    Dim Operators As Variant
    Operators = Array("", "xlBetween", "xlNotBetween", "xlEqual", _
                      "xlNotEqual", "xlGreater", "xlLess", "xlGreaterEqual", _
                      "xlLessEqual")
    
    Dim AlertStyle As Variant
    AlertStyle = Array("", "xlValidAlertStop", "xlValidAlertWarning", _
                       "xlValidAlertInformation")
    
    Dim RowCount As Long
    RowCount = 0
    Dim Tbl As TableManager.TableClass
    Dim I As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        RowCount = RowCount + Tbl.CellCount
    Next I
    
    Dim TAry As Variant
    ReDim TAry(RowCount, UBound(TitleArray))
    For I = 0 To UBound(TitleArray)
        TAry(0, I) = TitleArray(I)
    Next I
    
    RowCount = 1
    Dim Cll As TableManager.CellClass
    Dim J As Long
    For I = 0 To pAllTbls.Count - 1
        Set Tbl = Table(I, Module_Name)
        For J = 0 To Tbl.CellCount - 1
            Set Cll = Tbl.TableCells(J)
            TAry(RowCount, 0) = Tbl.Name
            TAry(RowCount, 1) = Cll.Name
            TAry(RowCount, 2) = Cll.HeaderText
            TAry(RowCount, 3) = CellTypes(Cll.CellType)
            If Cll.CellType <> xlValidateInputOnly Then
                If Cll.CellType <> xlValidateList Then
                    TAry(RowCount, 4) = Operators(Cll.Operator)
                End If
                TAry(RowCount, 5) = AlertStyle(Cll.ValidAlertStyle)
                If Left(Cll.ValidationFormula1, 1) = "=" Then
                    TAry(RowCount, 6) = "''" & Cll.ValidationFormula1
                Else
                    TAry(RowCount, 6) = Cll.ValidationFormula1
                End If
                TAry(RowCount, 7) = Cll.ValidationFormula2
            End If
            TAry(RowCount, 8) = IIf(Cll.IgnoreBlank, "True", "False")
            TAry(RowCount, 9) = Cll.ShowInput
            TAry(RowCount, 10) = Cll.InputTitle
            TAry(RowCount, 11) = Cll.InputMessage
            TAry(RowCount, 12) = Cll.ShowError
            TAry(RowCount, 13) = Cll.ErrorTitle
            TAry(RowCount, 14) = Cll.ErrorMessage
            RowCount = RowCount + 1
        Next J
    Next I
    
    BuildNewTable WkBkName, TAry, RowCount
End Sub

Private Sub BuildNewTable( _
        ByVal WkBkName As Variant, _
        ByVal TAry As Variant, _
        ByVal RowCount As Long)
    
    Dim Rng As String
    
    With Workbooks(WkBkName)
        If Not Contains(.Worksheets, "Parameters") Then
            .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).Name = "Parameters"
        End If
        With .Worksheets("Parameters")
            If Contains(.ListObjects, "ParameterTable") Then
                .ListObjects("ParameterTable").Delete
            End If
            .Range("$A$1").Resize(UBound(TAry, 1), UBound(TAry, 2)).Value = TAry
    
            Rng = "$A$1:" & ConvertToLetter(UBound(TAry, 2)) & RowCount - 1
            .ListObjects.Add(xlSrcRange, .Range(Rng), , xlYes).Name = "ParameterTable"
    
            .Activate
            .Range("A2").Select
        End With
    End With
    ActiveWindow.FreezePanes = True
    ActiveSheet.Range(Rng).EntireColumn.AutoFit
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


