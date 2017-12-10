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

Private Const Module_Name As String = "CellClass."

'Private pWorksheetName As String
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
Private pDBColNum As Long
Private pDBRange As Range

' Attributes for cells with validation
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
       ByVal Tbl As TableManager.TableClass, _
       ByVal CellID As Range)

    Dim DataCell As Range
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "CollectCellData"
    On Error GoTo ErrorHandler
    
    Debug.Assert CellID.Count = 1
    
    pHeaderText = CellID
    pName = Replace(pHeaderText, " ", vbNullString)
    pTableName = Tbl.Name
    
    Set DataCell = CellID.Offset(1, 0)
    pColumnWidth = CellID.ColumnWidth
    pDBColNum = Application.WorksheetFunction.Match(CellID, CellID.ListObject.HeaderRowRange, 0)
    pLocked = DataCell.Locked
    pNumberFormat = DataCell.NumberFormat
    pWrap = DataCell.WrapText
    Set pDBRange = CellID.ListObject.DataBodyRange
    
    If HasVal(DataCell) Then
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
        If pValidationFormula1 <> vbNullString Then
            pValidationList = DataCell.Parent.Evaluate(pValidationFormula1)
            If IsError(pValidationList) Then _
               pValidationList = Evaluate(pValidationFormula1)
            If IsError(pValidationList) Then _
               pValidationList = Split(pValidationFormula1, ",")
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' CollectCellData

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

Public Sub LetCellValue(ByVal RowNum As Long, ByVal Vbl As Variant)
    pDBRange(RowNum, pDBColNum) = Vbl
End Sub                                          ' LetCellValue

Public Property Get Length() As Long: Length = Len(pHeaderText): End Property

Public Property Get Control() As Control: Set Control = pControl: End Property

Public Property Set Control(ByVal Ctl As Control): Set pControl = Ctl: End Property

Public Property Get TableName() As String: TableName = pTableName: End Property

' Attributes for cells with validation
Public Property Get CellType() As XlDVType: CellType = pCellValidationType: End Property

Public Property Let CellType(ByVal CT As XlDVType): pCellValidationType = CT: End Property

Public Property Get Operator() As XlFormatConditionOperator: Operator = pOperator: End Property

Public Property Let Operator(ByVal Op As XlFormatConditionOperator): pOperator = Op: End Property

Public Property Get ValidAlertStyle() As XlDVAlertStyle: ValidAlertStyle = pValidAlertStyle: End Property

Public Property Let ValidAlertStyle(ByVal VAS As XlDVAlertStyle): pValidAlertStyle = VAS: End Property

Public Property Get ValidationFormula1() As String: ValidationFormula1 = pValidationFormula1: End Property

Public Property Let ValidationFormula1(ByVal VF1 As String): pValidationFormula1 = VF1: End Property

Public Property Get ValidationFormula2() As String: ValidationFormula2 = pValidationFormula2: End Property

Public Property Let ValidationFormula2(ByVal VF2 As String): pValidationFormula2 = VF2: End Property

Public Property Get IgnoreBlank() As Boolean: IgnoreBlank = pIgnoreBlank: End Property

Public Property Let IgnoreBlank(ByVal IB As Boolean): pIgnoreBlank = IB: End Property

Public Property Get ValidationList() As Variant: ValidationList = pValidationList: End Property

Public Property Let ValidationList(ByVal VL As Variant): pValidationList = VL: End Property

' Attributes for cells with validation inputs messages
Public Property Get ShowInput() As Boolean: ShowInput = pShowInput: End Property

Public Property Let ShowInput(ByVal SI As Boolean): pShowInput = SI: End Property

Public Property Get InputTitle() As String: InputTitle = pInputTitle: End Property

Public Property Let InputTitle(ByVal IT As String): pInputTitle = IT: End Property

Public Property Get InputMessage() As String: InputMessage = pInputMessage: End Property

Public Property Let InputMessage(ByVal IM As String): pInputMessage = IM: End Property

' Attributes for cells with error messages
Public Property Get ShowError() As Boolean: ShowError = pShowError: End Property

Public Property Let ShowError(ByVal SE As Boolean): pShowError = SE: End Property

Public Property Get ErrorTitle() As String: ErrorTitle = pErrorTitle: End Property

Public Property Let ErrorTitle(ByVal ET As String): pErrorTitle = ET: End Property

Public Property Get ErrorMessage() As String: ErrorMessage = pErrorMessage: End Property

Public Property Let ErrorMessage(ByVal EM As String): pErrorMessage = EM: End Property








