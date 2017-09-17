VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : CellClass
' Author : Owner
' Date   : 7/19/2017
' Purpose: Holds all data associated with an individual table cell/column
'---------------------------------------------------------------------------------------
Option Explicit

Private Const Module_Name = "CellClass."

Private pHeaderText As String
Private pName As String
Private pColumnWidth As Single
Private pLocked As Variant
Private pNumberFormat As Variant
Private pIgnoreBlank As Boolean
Private pValidAlertStyle As XlDVAlertStyle
Private pHasValidation As Boolean
Private pValidationType As XlDVType
Private pWrap As Boolean
Private pValidationFormula1 As String
Private pValidationFormula2 As String
Private pOperator As XlFormatConditionOperator
Private pValidationList As Variant
Private pFormControl As Control
Private pControlValue As String
Private pCellValue As String
Private pCell As Range
Private pDBColNum As Long
Private pDBRange As Range
Private pControl As TableManager.ControlClass

Private pShowInput As Boolean
Private pInputTitle As String
Private pInputMessage As String

Private pShowError As Boolean
Private pErrorTitle As String
Private pErrorMessage As String

'---------------------------------------------------------------------------------------
' Method : CollectData
' Author : Owner
' Date   : 7/19/2017
' Purpose: Collects all the data for one table cell/column
'---------------------------------------------------------------------------------------
Public Sub CollectData(ByVal CellID As Range)

'   Description: Description of what function does
'   Inputs:
'   CellID       The cell we're gathering the formats from
'   Outputs:
'   Me       Success/Failure
'   Requisites:
'   SharedRoutines
'   Notes:
'   Any notes
'   Example:
'   How to call this routine
'   History
'   05/14/2017 RRD Initial Programming

'   Declarations
    Const Routine_Name = Module_Name & "CollectData"
    Dim DataCell As Range

'   Error Handling Initialization
    On Error GoTo ErrHandler
    
'   Procedure
    Debug.Assert CellID.Count = 1
    
    pHeaderText = CellID
    
    Set DataCell = CellID.Offset(1, 0)
    Set pCell = CellID
    pColumnWidth = CellID.ColumnWidth
    pCellValue = CellID
    pDBColNum = Application.Match(CellID, CellID.ListObject.HeaderRowRange, 0)
    pName = Replace(pHeaderText, " ", "")
    pLocked = DataCell.Locked
    pNumberFormat = DataCell.NumberFormat
    pWrap = DataCell.WrapText
    Set pDBRange = CellID.ListObject.DataBodyRange
    
    pHasValidation = False
    If HasVal(DataCell) Then
        pHasValidation = True
        pValidationType = DataCell.Validation.Type
        pIgnoreBlank = DataCell.Validation.IgnoreBlank
        pValidAlertStyle = DataCell.Validation.AlertStyle
        
        pShowInput = DataCell.Validation.ShowInput
        pInputTitle = DataCell.Validation.InputTitle
        pInputMessage = DataCell.Validation.InputMessage
        
        pShowError = DataCell.Validation.ShowError
        pErrorTitle = DataCell.Validation.ErrorTitle
        pErrorMessage = DataCell.Validation.ErrorMessage
        
        pValidationType = DataCell.Validation.Type
        pValidationFormula1 = DataCell.Validation.Formula1
        pValidationFormula2 = DataCell.Validation.Formula2
        pOperator = DataCell.Validation.Operator
        
        Select Case pValidationType
        Case xlValidateList
            If pValidationFormula1 <> "" Then
                pValidationList = DataCell.Parent.Evaluate(pValidationFormula1)
                If IsError(pValidationList) Then pValidationList = Evaluate(pValidationFormula1)
                If IsError(pValidationList) Then pValidationList = Split(pValidationFormula1, ",")
            End If
        Case xlValidateWholeNumber
            ' Already collected Formula1, Formula2, Operator
        End Select
        
    End If
    
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Else: 'Display Error Message Box
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select

End Sub      ' CollectData

Public Function GetCellValue(ByVal RowNum As Long) As Variant
    
    GetCellValue = pDBRange(RowNum, pDBColNum)
    
End Function

Private Function HasVal(ByVal oTarget As Range) As Boolean

'   Description:Determine if a cell has validation
'   Inputs: oTarget     A single cell
'   Outputs: me          TRUE/FALSE
'   Requisites: *None
'   Example: ?HasVal(Selection)

'     Date   Ini Modification
'   04/16/2011 CWH Initial Programming
'   05/14/2017 RRD Moved to CellClass
'                   Made Private
'                   Changed name from HasValidation to HasVal
'                   Changed to Routine_Name

'   Declarations
    Const Routine_Name = Module_Name & "HasValidation"
    Dim v               As Variant
    
'   Error Handling Initialization
    On Error GoTo ErrHandler
    HasVal = False

'   Procedure
    v = oTarget.Validation.Type
    HasVal = True
        
ErrHandler:
    Select Case Err.Number
        Case Is = TableManager.NoError: 'Do nothing
        Case Is = 1004: 'Ignore: No Validation
        Case Else:
            Select Case TableManager.DspErrMsg(Routine_Name)
                Case Is = vbAbort: Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry: Resume          'Try again
                Case Is = vbIgnore: 'End routine
            End Select
    End Select

End Function

Public Sub LetCellValue(ByVal RowNum As Long, ByVal Vbl As Variant)
    pDBRange(RowNum, pDBColNum) = Vbl
End Sub

Public Property Get ColumnWidth() As Single: ColumnWidth = pColumnWidth: End Property

Public Property Get ControlValue() As String: ControlValue = pControlValue: End Property

Public Property Let ControlValue(ByVal Vbl As String): pControlValue = Vbl: End Property

Public Property Get Ctl() As TableManager.ControlClass: Set Ctl = pControl: End Property

Public Property Set Ctl(ByVal Ctl As TableManager.ControlClass): Set pControl = Ctl: End Property

Public Property Get ErrorMessage() As String: ErrorMessage = pErrorMessage: End Property

Public Property Get ErrorTitle() As String: ErrorTitle = pErrorTitle: End Property

Public Property Get FormControl() As Control: Set FormControl = pFormControl: End Property

Public Property Set FormControl(ByVal Vbl As Control): Set pFormControl = Vbl: End Property

Public Property Get HasValidation() As Boolean: HasValidation = pHasValidation: End Property

Public Property Let HasValidation(ByVal Vbl As Boolean): pHasValidation = Vbl: End Property

Public Property Get HeaderText() As String: HeaderText = pHeaderText: End Property

Public Property Get IgnoreBlank() As Boolean: IgnoreBlank = pIgnoreBlank: End Property

Public Property Get InputMessage() As String: InputMessage = pInputMessage: End Property

Public Property Get InputTitle() As String: InputTitle = pInputTitle: End Property

Public Property Get Length() As Integer: Length = Len(pHeaderText): End Property

Public Property Get Locked() As Variant: Locked = pLocked: End Property

Public Property Get Name() As String: Name = pName: End Property

Public Property Get NumberFormat() As Variant: NumberFormat = pNumberFormat: End Property

Public Property Get ShowError() As Boolean: ShowError = pShowError: End Property

Public Property Get ShowInput() As Boolean: ShowInput = pShowInput: End Property

Public Property Get ValidAlertStyle() As XlDVAlertStyle: ValidAlertStyle = pValidAlertStyle: End Property

Public Property Get ValidationFormula1() As String: ValidationFormula1 = pValidationFormula1: End Property

Public Property Get ValidationFormula2() As String: ValidationFormula2 = pValidationFormula2: End Property

Public Property Get ValidationList() As Variant: ValidationList = pValidationList: End Property

Public Property Get ValidationOperator() As String: ValidationOperator = pOperator: End Property

Public Property Get ValidationType() As XlDVType: ValidationType = pValidationType: End Property

Public Property Let ValidationType(ByVal Vbl As XlDVType): pValidationType = Vbl: End Property

Public Property Get Wrap() As Boolean: Wrap = pWrap: End Property
