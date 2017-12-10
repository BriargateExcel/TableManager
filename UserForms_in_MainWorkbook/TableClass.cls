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
End Sub ' Class_Initialize

Public Property Get ActiveRow() As Long
    ActiveRow = Me.DBRow
End Property

Public Property Get ActiveTarget() As Range: Set ActiveTarget = pTarget: End Property
Public Property Set ActiveTarget(ByVal Target As Range)
    Target.Select
    Set pTarget = Target
End Property

Public Property Get DBRange() As Range
    Set DBRange = MainWorkbook.Worksheets(pWkShtName).ListObjects(pTableName).DataBodyRange
End Property

Public Property Get Form() As TableManager.FormClass: Set Form = pForm: End Property
Public Property Set Form(ByVal Frm As TableManager.FormClass): Set pForm = Frm: End Property

Public Property Get Locked() As Boolean: Locked = pLocked: End Property

Public Property Get MaxFieldLength() As Single: MaxFieldLength = pMaxField: End Property

Public Property Get MaxHeaderLength() As Single: MaxHeaderLength = pMaxHeader: End Property

Public Property Get Name() As String: Name = pTableName: End Property
Public Property Let Name(ByVal TN As String): pTableName = TN: End Property

Public Property Get NumColumns() As Long: NumColumns = pCells.Count: End Property

Public Property Get NumRows() As Long: NumRows = DBRange.Rows.Count: End Property

Public Property Get TableCells() As TableManager.CellsClass: Set TableCells = pCells: End Property

Public Property Get Target() As Range: Set Target = pTarget: End Property

Public Property Set Target(ByVal Vbl As Range): Set pTarget = Vbl: End Property

Public Property Get WorksheetName() As String: WorksheetName = pWkShtName: End Property

Public Property Get WrapCount() As Long: WrapCount = pWrapCount: End Property

Public Property Get Table() As ListObject: Set Table = pTable: End Property
Public Property Set Table(ByVal Tbl As ListObject): Set pTable = Tbl: End Property

'@Ignore FunctionReturnValueNotUsed
Public Function CollectTableData( _
    ByVal WS As TableManager.WorksheetClass, _
    ByVal Tbl As TableManager.TableClass _
    ) As Boolean

    Dim Element As Range
    Dim OneCell As TableManager.CellClass

    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "CollectTableData"
    On Error GoTo ErrorHandler
    
    pTableName = Tbl.Name
    pWkShtName = WS.Worksheet.Name
    pLocked = True
    
    For Each Element In pTable.HeaderRowRange
        Set OneCell = New TableManager.CellClass
        OneCell.CollectCellData Tbl, Element
        pCells.Add OneCell
        If OneCell.ColumnWidth > pMaxField Then pMaxField = OneCell.ColumnWidth
        If OneCell.Wrap Then pWrapCount = pWrapCount + 1
        If OneCell.Length > pMaxHeader Then pMaxHeader = OneCell.Length
        If Not OneCell.Locked Then pLocked = False
    Next Element
    
    If pLocked Then
        CollectTableData = False
    Else
        CollectTableData = True
    End If

'@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' CollectTableData

Public Function FirstCell() As Range
    Set FirstCell = pTable.DataBodyRange(1, 1)
End Function ' FirstCell

Public Function LastCell() As Range
'    Set LastCell = pTable.DataBodyRange(pTable.ListRows.Count, pTable.ListColumns.Count)
    Set LastCell = pTable.DataBodyRange(pTable.ListRows.Count, 1)
End Function ' LastCell

Public Function LastRow() As Long
    LastRow = pTarget.ListObject.ListRows.Count
End Function ' LastRow

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
End Sub ' NextRow

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
End Sub ' PreviousRow

Public Function DBRow() As Long
    Dim THRow As Long: THRow = Me.ActiveTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = Me.ActiveTarget.Row
    DBRow = TRow - THRow
End Function ' DBRow

Public Function DBCol(ByVal ColumnName As String) As Long
    Dim Rng As Range
    Dim v As Variant
    
    Set Rng = Me.ActiveTarget.ListObject.HeaderRowRange
    v = Application.WorksheetFunction.Match(ColumnName, Rng, 0)
    DBCol = v
End Function ' DBCol

Public Function CellCount() As Long: CellCount = pCells.Count: End Function





