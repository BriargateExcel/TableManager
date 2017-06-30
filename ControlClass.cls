VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "ControlClass."

Private pHeaderText As String
Private pColumnWidth As Single
Private pLocked As Variant
Private pNumberFormat As Variant
Private pCellType As XlDVType   ' https://msdn.microsoft.com/en-us/library/office/ff840715.aspx
Private pIgnoreBlank As Boolean
Private pValidAlertStyle As XlDVAlertStyle ' https://translate.google.com/translate?hl=en&sl=ja&u=https://msdn.microsoft.com/ja-jp/library/office/ff841223.aspx&prev=search
Private pHasValidation As Boolean
Private pWrap As Boolean
Private pValidationFormula As String
Private pValidationList As Variant

Private pShowError As Boolean
Private pErrorMessage As String
Private pErrorTitle As String

Private pShowInput As Boolean
Private pInputMessage As String
Private pInputTitle As String


Public Function SetHeader(ByVal CellID As Range) As Boolean

'   Description: Gathers the data about one column header in a table
'   Inputs:
'   CellID       The cell we're gathering the data from
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
    Const Routine_Name = Module_Name & "SetHeader"

'   Error Handling Initialization
    On Error GoTo ErrHandler
    SetHeader = Failure
    
'   Procedure
    Debug.Assert CellID.Count = 1
    
    pHeaderText = CellID.Value
    
    SetHeader = Success
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

End Function ' SetHeader

Private Function HasVal(ByVal oTarget As Range) As Boolean

'   Description:Determine if a cell has validation
'   Inputs:     oTarget     A single cell
'   Outputs:    me          TRUE/FALSE
'   Requisites: *None
'   Example:    ?HasVal(Selection)

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
    pHasValidation = True
    HasVal = True
        
ErrHandler:
    Select Case Err.Number
        Case Is = NoError:                          'Do nothing
        Case Is = 1004:                             'Ignore: No Validation
        Case Else:
            Select Case DspErrMsg(Routine_Name)
                Case Is = vbAbort:  Stop: Resume    'Debug mode - Trace
                Case Is = vbRetry:  Resume          'Try again
                Case Is = vbIgnore:                 'End routine
            End Select
    End Select

End Function


Public Property Get ColumnWidth() As Single: ColumnWidth = pColumnWidth: End Property

Public Property Get Locked() As Variant: Locked = pLocked: End Property

Public Property Get NumberFormat() As Variant: NumberFormat = pNumberFormat: End Property

Public Property Get CellType() As XlDVType: CellType = pCellType: End Property

Public Property Get ValidAlertStyle() As XlDVAlertStyle: ValidAlertStyle = pValidAlertStyle: End Property

Public Property Get ShowError() As Boolean: ShowError = pShowError: End Property

Public Property Get ErrorMessage() As String: ErrorMessage = pErrorMessage: End Property

Public Property Get ErrorTitle() As String: ErrorTitle = pErrorTitle: End Property

Public Property Get ShowInput() As Boolean: ShowInput = pShowInput: End Property

Public Property Get InputMessage() As String: InputMessage = pInputMessage: End Property

Public Property Get InputTitle() As String: InputTitle = pInputTitle: End Property

Public Property Get IgnoreBlank() As Boolean: IgnoreBlank = pIgnoreBlank: End Property

Public Property Get HeaderText() As String: HeaderText = pHeaderText: End Property

Public Property Get Wrap() As Boolean: Wrap = pWrap: End Property

Public Property Get Length() As Integer: Length = Len(pHeaderText): End Property

Public Property Get ValidationFormula() As String: ValidationFormula = pValidationFormula: End Property

Public Property Get ValidationList() As Variant: ValidationList = pValidationList: End Property

