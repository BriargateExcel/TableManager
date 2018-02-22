VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name As String = "TableClass."

Private pTableName As String
Private pWkShtName As String
Private pWkSht As Worksheet
Private pForm As TableManager.FormClass
Private pCells As TableManager.CellsClass
Private pTable As ListObject

Private pMaxHeader As Single
Private pMaxField As Single
Private pWrapCount As Single
Private pLocked As Boolean
Private pTarget As Range

Private Sub Class_Initialize()
    Set pCells = New TableManager.CellsClass
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("TableRoutines.")
End Function                                     ' ModuleList

Public Property Get ActiveRow() As Long
    ActiveRow = Me.DBRow
End Property

Public Property Get ActiveTarget() As Range: Set ActiveTarget = pTarget: End Property
Public Property Set ActiveTarget(ByVal Target As Range)
    Target.Select
    Set pTarget = Target
End Property

Public Sub GoToHome()
    Set pTarget = Me.FirstCell
End Sub

Public Property Get DBRange() As Range
    Set DBRange = pWkSht.ListObjects(pTableName).DataBodyRange
End Property

Public Property Get DBColRange( _
       ByVal Tbl As TableManager.TableClass, _
       ByVal ColumnName As String _
       ) As Range
    
    Dim ColNum As Long
    ColNum = Tbl.DBColNumber(ColumnName)
    If ColNum = 0 Then
        Err.Raise 1, "TableClass.DBColRange", "Fatal error. ColumnName not found."
    End If
    
    Set DBColRange = Tbl.Table.ListColumns(ColNum).DataBodyRange
End Property

Public Property Get DBColNumber(ByVal ColumnName As String) As Long
    On Error Resume Next
    DBColNumber = Application.WorksheetFunction.Match(ColumnName, Me.Table.HeaderRowRange, 0)
    If Err.Number <> 0 Then DBColNumber = 0
    On Error GoTo 0
End Property

Public Property Get DBRowNumber( _
       ByVal ColumnName As String, _
       ByVal Key As String _
       ) As Long
    
    Dim KeyRange As Range
    Set KeyRange = Me.DBColRange(Me, ColumnName)
    
    On Error Resume Next
    DBRowNumber = Application.WorksheetFunction.Match(Key, KeyRange, 0)
    If Err.Number <> 0 Then DBRowNumber = 0
    On Error GoTo 0
End Property

Public Property Get Form() As TableManager.FormClass: Set Form = pForm: End Property
Public Property Set Form(ByVal Frm As TableManager.FormClass): Set pForm = Frm: End Property

Public Property Get Locked() As Boolean: Locked = pLocked: End Property

Public Property Get MaxFieldLength() As Single: MaxFieldLength = pMaxField: End Property

Public Property Get MaxHeaderLength() As Single: MaxHeaderLength = pMaxHeader: End Property

Public Property Get Name() As String: Name = pTableName: End Property

Public Property Let Name(ByVal Vbl As String): pTableName = Vbl: End Property

Public Property Get NumColumns() As Long: NumColumns = pCells.Count: End Property

Public Property Get NumRows() As Long: NumRows = DBRange.Rows.Count: End Property

Public Property Get TableCells() As TableManager.CellsClass: Set TableCells = pCells: End Property

Public Property Get Target() As Range: Set Target = pTarget: End Property

Public Property Set Target(ByVal Vbl As Range): Set pTarget = Vbl: End Property

Public Property Get WorksheetName() As String: WorksheetName = pWkShtName: End Property

Public Property Get Worksheet() As Worksheet: Set Worksheet = pWkSht: End Property

Public Property Get WrapCount() As Long: WrapCount = pWrapCount: End Property

Public Property Get Table() As ListObject: Set Table = pTable: End Property

Public Property Set Table(ByVal Tbl As ListObject): Set pTable = Tbl: End Property

'@Ignore FunctionReturnValueNotUsed
Public Function CollectTableData( _
       ByVal Tbl As TableManager.TableClass, _
       ByVal Modulename As String _
       ) As Boolean
    
    Const RoutineName As String = Module_Name & "CollectTableData"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.Initializing

    Debug.Assert InScope(ModuleList, Modulename)
    
    pTableName = Tbl.Table.Name
    pWkShtName = Tbl.Table.Parent.Name
    Set pWkSht = TableManager.GetMainWorkbook.Worksheets(pWkShtName)
    pLocked = True
    Dim Element As Range
    Dim OneCell As TableManager.CellClass
    
    For Each Element In pTable.HeaderRowRange
        Set OneCell = New TableManager.CellClass
        OneCell.CollectCellData Tbl, Element, Module_Name
        pCells.Add OneCell, Module_Name
        If OneCell.ColumnWidth > pMaxField Then pMaxField = OneCell.ColumnWidth
        If OneCell.Wrap Then pWrapCount = pWrapCount + 1
        If OneCell.Length > pMaxHeader Then pMaxHeader = OneCell.Length
        If Not OneCell.Locked Then pLocked = False
    Next Element
    

    CollectTableData = Not pLocked

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' CollectTableData

Public Function FirstCell() As Range
    Const RoutineName As String = Module_Name & "FirstCell"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set FirstCell = pTable.DataBodyRange(1, 1)
    If Err.Number = 91 Then
        Set FirstCell = pTable.HeaderRowRange(2, 1)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' FirstCell

Public Function LastCell() As Range
    Const RoutineName As String = Module_Name & "LastCell"
    On Error GoTo ErrorHandler
    
    Set LastCell = pTable.DataBodyRange(pTable.ListRows.Count, 1)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' LastCell

Public Function LastRow() As Long
    LastRow = pTarget.ListObject.ListRows.Count
End Function                                     ' LastRow

Public Sub NextRow()
    Dim DBRow As Long
    Dim I As Long
    
    DBRow = Me.DBRow
    
    If DBRow < pTarget.ListObject.ListRows.Count Then
        TableManager.TurnOnCellDescriptions Me, Module_Name
        For I = DBRow To DBRange.Rows.Count
            Set pTarget = pTarget.Offset(1)
            If Not pTarget.Rows.Hidden Then Exit For
        Next I
        TableManager.TurnOffCellDescriptions Me, Module_Name
    Else
        MsgBox "Cannot advance to the next row; already at last row", _
               vbOKOnly Or vbExclamation, "Already at End of Table"
    End If
End Sub                                          ' NextRow

Public Sub PreviousRow()
    Dim DBRow As Long
    Dim I As Long
    
    DBRow = Me.DBRow
    
    If DBRow > 1 Then
        TableManager.TurnOnCellDescriptions Me, Module_Name
        For I = DBRow To DBRange.Rows.Count
            Set pTarget = pTarget.Offset(-1)
            If Not pTarget.Rows.Hidden Then Exit For
        Next I
        TableManager.TurnOffCellDescriptions Me, Module_Name
    Else
        MsgBox "Cannot move to the previous row; already at first row", _
               vbOKOnly Or vbExclamation, _
               "Already at Beginning of Table"
    End If
End Sub                                          ' PreviousRow

Public Function DBRow() As Long
    Dim THRow As Long: THRow = Me.ActiveTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = Me.ActiveTarget.Row
    DBRow = TRow - THRow
End Function                                     ' DBRow

Public Function SelectedDBCol(ByVal ColumnName As String) As Long
    Dim Rng As Range
    Dim Vbl As Variant
    
    Set Rng = Me.ActiveTarget.ListObject.HeaderRowRange
    On Error Resume Next
    Vbl = Application.WorksheetFunction.Match(ColumnName, Rng, 0)
    If Err.Number <> 0 Then
        SelectedDBCol = 0
        Exit Function
    End If
    On Error GoTo 0
    SelectedDBCol = Vbl
End Function                                     ' SelectedDBCol

Public Function ColumnRange(ByVal ColNum As Long) As Range
    Set ColumnRange = pWkSht.Range(DBRange(1, ColNum), DBRange(NumRows, ColNum))
End Function

Public Function CellCount() As Long: CellCount = pCells.Count: End Function


