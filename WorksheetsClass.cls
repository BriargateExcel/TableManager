VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Worksheets")

Option Explicit

Private Const Module_Name As String = "WorksheetsClass."

Private pAllSheets As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllSheets = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Public Property Get Count() As Long: Count = TableCount(pAllSheets.Count): End Property

'@DefaultMember
Public Property Get Item(ByVal vIndex As Variant) As WorksheetClass
    'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName As String = Module_Name & "Get Item"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set Item = pAllSheets.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllSheets(vIndex)
    On Error GoTo 0
    
    Debug.Assert Not Item Is Nothing
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Property
    '@Ignore LineLabelNotUsed
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Sub Add(ByVal Sht As WorksheetClass)

    Dim Evt As EventClass
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Add"
    On Error GoTo ErrorHandler
    
    Set Evt = New EventClass
    
    If Sht.Name <> vbNullString Then
        pAllSheets.Add Sht.Name, Sht
        Set Evt.SheetEvent = Sht.Worksheet
        Set pAllSheets.Item(Sht.Worksheet.Name).WorksheetEvent.SheetEvent = Sht.Worksheet
    End If
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' Add

Public Sub Remove(ByVal vIndex As Variant)

    Debug.Assert Initializing
    
    Const RoutineName As String = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        TableSetNothing Module_Name
        TableSetNewDict Module_Name
    Else
        If Not TableExists(vIndex, Module_Name) Then Err.Raise 9
        TableRemove vIndex, Module_Name
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Remove

