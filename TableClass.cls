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

Private Const Module_Name = "TableClass."

Private pCells As CellsClass

Private pMaxHeader As Single
Private pMaxField As Single
Private pWrapCount As Single
Private pLocked As Boolean
Private pName As String
Private pTable As ListObject
Private pTarget As Range
Private pForm As FormClass
Private pWkSht As Worksheet

Public Property Get ActiveRow() As Long
    Dim THRow As Long: THRow = pTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = pTarget.Row
    Dim DBrow As Long: DBrow = TRow - THRow

    ActiveRow = DBrow

End Property

Public Property Get ActiveTarget() As Range: Set ActiveTarget = pTarget: End Property

Public Property Set ActiveTarget(ByVal Vbl As Range)
    Vbl.Select
    Set pTarget = Vbl
End Property

Public Property Get DBRange() As Range
    Set DBRange = pWkSht.ListObjects(pName).DataBodyRange
End Property

Public Property Get Form() As FormClass:          Set Form = pForm:             End Property

Public Property Set Form(ByVal Vbl As FormClass): Set pForm = Vbl:              End Property

Public Property Get Locked() As Boolean:          Locked = pLocked:             End Property

Public Property Get MaxFieldLength() As Single:   MaxFieldLength = pMaxField:   End Property

Public Property Get MaxHeaderLength() As Single:  MaxHeaderLength = pMaxHeader: End Property

Public Property Get Name() As String:             Name = pName:                 End Property

Public Property Get NumFields() As Integer:       NumFields = pCells.Count:     End Property

Public Property Get TableCells() As CellsClass:   Set TableCells = pCells:      End Property

Public Property Get Target() As Range:            Set Target = pTarget:         End Property

Public Property Set Target(ByVal Vbl As Range):   Set pTarget = Vbl:            End Property

Public Property Get Worksheet() As Worksheet:     Set Worksheet = pWkSht:       End Property

Public Property Get WrapCount() As Integer:       WrapCount = pWrapCount:       End Property

Public Sub CollectData( _
    ByVal WS As Worksheet, _
    ByVal TblName As String)

'   Description: Description of what function does
'   Inputs:
'   WS          Worksheet containing the table
'   TblName     Name of the table under analysis
'   Outputs:
'   Me       Success/Failure
'   Requisites:
'   None
'   Notes:
'   Any notes
'   Example:
'   How to call this routine
'   History
'   05/14/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "CollectData"
    Dim Element As Range
    Dim OneCell As CellClass

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    pName = TblName
    Set pWkSht = WS
    Set pTable = WS.ListObjects(TblName)
    Set pCells = New CellsClass
    pLocked = True
    
    For Each Element In pTable.HeaderRowRange
        Set OneCell = New CellClass
        OneCell.CollectData Element
        pCells.Add OneCell
        If OneCell.ColumnWidth > pMaxField Then pMaxField = OneCell.ColumnWidth
        If OneCell.Wrap Then pWrapCount = pWrapCount + 1
        If OneCell.Length > pMaxHeader Then pMaxHeader = OneCell.Length
        If Not OneCell.Locked Then pLocked = False
    Next Element
    
    If pLocked Then
        MsgBox "There are no editable fields in this table", _
            vbOKOnly Or vbExclamation, "No Editable Fields"
        Stop
        ' ToDo: Finalize this so it returns appropriately
    End If

ErrHandler:
    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Else:
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Sub      ' CollectData

Public Function FirstCell() As Range
    Set FirstCell = pTable.DataBodyRange(1, 1)
End Function

Public Function LastCell() As Range
    Set LastCell = pTable.DataBodyRange(pTable.ListRows.Count, pTable.ListColumns.Count)
End Function

Public Function LastRow() As Long
    LastRow = pTarget.ListObject.ListRows.Count
End Function

Public Sub NextRow()
    Dim THRow As Long: THRow = pTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = pTarget.Row
    Dim DBrow As Long: DBrow = TRow - THRow
    
    If DBrow < pTarget.ListObject.ListRows.Count Then
        Set pTarget = pTarget.Offset(1)
    Else
        MsgBox "Cannot advance to the next row; already at last row", vbOKOnly Or vbExclamation, "Already at End of Table"
    End If
End Sub

Public Sub PopulateTable(ByVal Target As Range)
    Dim Cll As CellClass
    Dim TRow As Long: TRow = Target.Row
    Dim THRow As Long: THRow = Target.ListObject.HeaderRowRange.Row
    Dim RowNum As Long: RowNum = TRow - THRow
    
    For Each Cll In pCells
        Cll.LetCellValue RowNum, Cll.FormControl
    Next Cll
    
End Sub

Public Sub PreviousRow()
    Dim THRow As Long: THRow = pTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = pTarget.Row
    Dim DBrow As Long: DBrow = TRow - THRow
    
    If DBrow > 1 Then
        Set pTarget = pTarget.Offset(-1)
    Else
        MsgBox "Cannot move to the previous row; already at first row", _
            vbOKOnly Or vbExclamation, _
            "Already at Beginning of Table"
    End If
End Sub

