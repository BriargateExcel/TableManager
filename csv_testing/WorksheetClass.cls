VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pWorksheet As Worksheet
Private pName As String

Private pWorksheetEvent As TableManager.EventClass

Private Sub Class_Initialize()
    Debug.Assert Initializing
    Set pWorksheetEvent = New TableManager.EventClass
End Sub                                          ' Class_Initialize

Public Property Get Worksheet() As Worksheet: Set Worksheet = pWorksheet: End Property
Public Property Set Worksheet(ByVal WS As Worksheet): Set pWorksheet = WS: End Property

Public Property Get WorksheetEvent() As Object: Set WorksheetEvent = pWorksheetEvent: End Property
Public Property Set WorksheetEvent(ByVal WS As Worksheet): Set pWorksheetEvent = WS: End Property

Public Property Get Name() As String: Name = pName: End Property
Public Property Let Name(ByVal NM As String): pName = NM: End Property

