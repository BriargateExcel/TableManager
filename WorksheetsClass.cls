VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "WorksheetsClass."

Private pAllSheets As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllSheets = New Scripting.Dictionary
End Sub ' Class_Initialize

Public Property Get Count() As Integer: Count = TableCount(pAllSheets.Count): End Property

Public Property Get Item(ByVal vIndex As Variant) As TableManager.WorksheetClass
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllSheets.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    Set Item = pAllSheets(vIndex)
    On Error GoTo 0
    
Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Sub Add(ByVal Sht As TableManager.WorksheetClass)

    Dim Evt As TableManager.EventClass
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Add"
    On Error GoTo ErrorHandler
    
    Set Evt = New TableManager.EventClass
    
    If Sht.Name <> "" Then
        pAllSheets.Add Sht.Name, Sht
        Set Evt.SheetEvent = Sht.Worksheet
        Set pAllSheets.Item(Sht.Worksheet.Name).WorksheetEvent.SheetEvent = Sht.Worksheet
    End If
        
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' Add

Public Function Exists(ByVal vIndex) As Boolean
    Exists = pAllSheets.Exists(vIndex)
End Function ' Exists

Public Function ListSheets() As Variant: ListSheets = pAllSheets.Keys: End Function

Public Sub Remove(ByVal vIndex As Variant)

    Debug.Assert Initializing
    
    Const RoutineName = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        TableSetNothing Module_Name
        TableSetNewDict Module_Name
    Else
        If Not TableExists(vIndex, Module_Name) Then Err.Raise 9
        TableRemove vIndex, Module_Name
    End If

Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Remove
