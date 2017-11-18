VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "TablesClass."

Private pAllTables As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllTables = New Scripting.Dictionary
End Sub ' Class_Initialize

Public Property Get Count() As Integer: Count = pAllTables.Count: End Property

Public Property Get Item(ByVal vIndex As Variant) As TableManager.TableClass
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0

    Const RoutineName = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllTables.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    Set Item = pAllTables(vIndex)
    On Error GoTo 0

Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Sub Add(ByVal Tbl As TableManager.TableClass)
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Add"
    On Error GoTo ErrorHandler
           
    If Tbl.Name <> "" Then pAllTables.Add Tbl.Name, Tbl
        
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Add

Public Function Exists(ByVal vIndex) As Boolean
    Exists = pAllTables.Exists(vIndex)
End Function ' Exists

Public Function ListTables() As Variant: Set ListTables = pAllTables.Keys: End Function

Public Sub Remove(ByVal vIndex As Variant)
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllTables = Nothing
        Set pAllTables = New Scripting.Dictionary
    Else
        If Not pAllTables.Exists(vIndex) Then Err.Raise 9
        pAllTables.Remove vIndex
    End If

Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Remove
