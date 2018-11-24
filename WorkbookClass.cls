VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("TableManager.Workbooks")

Option Explicit

Private Const Module_Name As String = "WorkbookClass."

Private pWorkbook As Workbook
Private pName As String
Private pTables As TablesClass

Private pWorkbookEvent As EventClass

Private Sub Class_Initialize()
    Debug.Assert Initializing
    Set pWorkbookEvent = New EventClass
    Set pTables = New TablesClass
End Sub                                          ' Class_Initialize

Public Property Get Workbook() As Workbook
    Set Workbook = pWorkbook
End Property

Public Property Set Workbook(ByVal WB As Worksheet)
    Set pWorkbook = WB
End Property

Public Property Get WorkbookEvent() As Object
    Set WorkbookEvent = pWorkbookEvent
End Property

Public Property Set WorkbookEvent(ByVal WB As Workbook)
    Set pWorkbookEvent = WB
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal NM As String)
    pName = NM
End Property

