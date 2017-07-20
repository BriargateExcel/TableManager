VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const cModule As String = "WorksheetClass"

'   Public Properties
Public WithEvents Worksheet As Worksheet        'Worksheet containing ComboBox
Attribute Worksheet.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set Worksheet = Nothing
End Sub


Private Sub Worksheet_BeforeDoubleClick( _
    ByVal Target As Range, _
    Cancel As Boolean)
    
    Dim Rng As Range
    Dim TableName As String
    Dim WS As Worksheet
    
    Cancel = True

    TableName = "TestTable"
    
    Set WS = Worksheets("Test Sheet")

    Set Rng = WS.ListObjects(TableName).DataBodyRange
    
    If Intersect(Target, Rng) Is Nothing Then
        MsgBox "The cursor must be inside a table", _
            vbOKOnly Or vbExclamation, _
            "Cursor Not in Table"
        Exit Sub
    End If
    
    If Target.Cells.Count > 1 Then
        MsgBox "Select only one cell inside the table before double clicking", _
            vbOKOnly Or vbExclamation, _
            "Multiple Cells Selected"
        Exit Sub
    End If
    
    BuildTable WS, TableName, Target
End Sub


