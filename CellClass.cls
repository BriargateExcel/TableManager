VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "CellClass."

Private pWorksheetName As String
Private pTableName As String
Private pControl As Control

' Attributes common to all table cells
Private pHeaderText As String
Private pName As String
Private pColumnWidth As Single
Private pLocked As Variant
Private pNumberFormat As Variant
Private pWrap As Boolean
Private pFormControl As Control
Private pControlValue As String
Private pCellValue As String
Private pCell As Range
Private pDBColNum As Long
Private pDBRange As Range

' Attributes for cells with validation
Private pHasValidation As Boolean
Private pCellValidationType As XlDVType
Private pOperator As XlFormatConditionOperator
Private pValidAlertStyle As XlDVAlertStyle
Private pValidationFormula1 As String
Private pValidationFormula2 As String
Private pIgnoreBlank As Boolean
Private pValidationList As Variant

' Attributes for cells with validation inputs messages
Private pShowInput As Boolean
Private pInputTitle As String
Private pInputMessage As String

' Attributes for cells with error messages
Private pShowError As Boolean
Private pErrorTitle As String
Private pErrorMessage As String

Public Sub CollectCellData( _
    ByVal WS As TableManager.WorksheetClass, _
    ByVal Tbl As TableManager.TableClass, _
    ByVal CellID As Range)

    Dim DataCell As Range
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "CollectCellData"
    On Error GoTo ErrorHandler
    
    Debug.Assert CellID.Count = 1
    
    pHeaderText = CellID
    pName = Replace(pHeaderText, " ", "")
    pWorksheetName = WS.Worksheet.Name
    pTableName = Tbl.Name
    
    Set DataCell = CellID.Offset(1, 0)
    Set pCell = CellID
    pColumnWidth = CellID.ColumnWidth
    pCellValue = CellID
    pDBColNum = Application.Match(CellID, CellID.ListObject.HeaderRowRange, 0)
    pLocked = DataCell.Locked
    pNumberFormat = DataCell.NumberFormat
    pWrap = DataCell.WrapText
    Set pDBRange = CellID.ListObject.DataBodyRange
    
    pHasValidation = False
    If HasVal(DataCell) Then
        pHasValidation = True
        pCellValidationType = DataCell.Validation.Type
        pIgnoreBlank = DataCell.Validation.IgnoreBlank
        pValidAlertStyle = DataCell.Validation.AlertStyle
        pOperator = DataCell.Validation.Operator
        
        pShowInput = DataCell.Validation.ShowInput
        pInputTitle = DataCell.Validation.InputTitle
        pInputMessage = DataCell.Validation.InputMessage
        
        pShowError = DataCell.Validation.ShowError
        pErrorTitle = DataCell.Validation.ErrorTitle
        pErrorMessage = DataCell.Validation.ErrorMessage
        
        pValidationFormula1 = DataCell.Validation.Formula1
        pValidationFormula2 = DataCell.Validation.Formula2
        If pValidationFormula1 <> "" Then
            pValidationList = DataCell.Parent.Evaluate(pValidationFormula1)
            If IsError(pValidationList) Then _
                pValidationList = Evaluate(pValidationFormula1)
            If IsError(pValidationList) Then _
                pValidationList = Split(pValidationFormula1, ",")
        End If
    End If
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' CollectCellData

Private Function HasVal(ByVal Target As Range) As Boolean

    Const RoutineName = Module_Name & "HasVal"
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    Debug.Assert Initializing
    
    On Error Resume Next
    
    v = Target.Validation.Type
    If Err.Number = 0 Then
        HasVal = True
    Else
        HasVal = False
    End If

Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' HasVal

' Attributes common to all table cells
Public Property Get HeaderText() As String: HeaderText = pHeaderText: End Property

Public Property Get Name() As String: Name = pName: End Property

Public Property Get ColumnWidth() As Single: ColumnWidth = pColumnWidth: End Property

Public Property Get Locked() As Variant: Locked = pLocked: End Property

Public Property Get NumberFormat() As Variant: NumberFormat = pNumberFormat: End Property

Public Property Get Wrap() As Boolean: Wrap = pWrap: End Property

Public Property Get FormControl() As Control: Set FormControl = pFormControl: End Property
Public Property Set FormControl(ByVal Vbl As Control): Set pFormControl = Vbl: End Property

Public Property Get ControlValue() As String: ControlValue = pControlValue: End Property
Public Property Let ControlValue(ByVal Vbl As String): pControlValue = Vbl: End Property

Public Function GetCellValue(ByVal RowNum As Long) As Variant
    GetCellValue = pDBRange(RowNum, pDBColNum)
End Function ' GetCellValue

Public Sub LetCellValue(ByVal RowNum As Long, ByVal Vbl As Variant)
    pDBRange(RowNum, pDBColNum) = Vbl
End Sub ' LetCellValue

Public Property Get Length() As Integer: Length = Len(pHeaderText): End Property

Public Property Get Control() As Control: Control = pControl: End Property
Public Property Let Control(ByVal Ctl As Control): pControl = Ctl: End Property

' Attributes for cells with validation
Public Property Get CellType() As XlDVType: CellType = pCellValidationType: End Property

Public Property Get Operator() As XlFormatConditionOperator: Operator = pOperator: End Property

Public Property Get ValidAlertStyle() As XlDVAlertStyle: ValidAlertStyle = pValidAlertStyle: End Property

Public Property Get ValidationFormula1() As String: ValidationFormula1 = pValidationFormula1: End Property

Public Property Get ValidationFormula2() As String: ValidationFormula2 = pValidationFormula2: End Property

Public Property Get IgnoreBlank() As Boolean: IgnoreBlank = pIgnoreBlank: End Property

Public Property Get ValidationList() As Variant: ValidationList = pValidationList: End Property

' Attributes for cells with validation inputs messages
Public Property Get ShowInput() As Boolean: ShowInput = pShowInput: End Property
Public Property Let ShowInput(ByVal Vis As Boolean): pShowInput = Vis: End Property

Public Property Get InputTitle() As String: InputTitle = pInputTitle: End Property

Public Property Get InputMessage() As String: InputMessage = pInputMessage: End Property

' Attributes for cells with error messages
Public Property Get ShowError() As Boolean: ShowError = pShowError: End Property

Public Property Get ErrorTitle() As String: ErrorTitle = pErrorTitle: End Property

Public Property Get ErrorMessage() As String: ErrorMessage = pErrorMessage: End Property

Public Property Get TableName() As String: TableName = pTableName: End Property
