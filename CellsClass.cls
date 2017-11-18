VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name = "CellsClass."

Private pAllCells As Scripting.Dictionary

' TODO:
' All the plural classes are built on a Dictionary
' The plural classes: Cells, Controls, Events, Tables, Worksheets
' They all have these routines: Count, Item, Exists, Add, Remove, Init, and List
' Is there any advantage in building them as Interface/Implements classes?
' CellsClass      is only used in TableClass
' ControlsClass   is only used in FormClass
' EventsClass     is only used in FormClass
' TablesClass     is only used in TableRoutines
' WorksheetsClass is only used in WorksheetRoutines

Private Sub Class_Initialize()
    Set pAllCells = New Scripting.Dictionary
End Sub ' Class_Initialize

Public Property Get Count() As Integer: Count = pAllCells.Count: End Property

Public Property Get Item(ByVal vIndex As Variant) As TableManager.CellClass
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllCells.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    Set Item = pAllCells(vIndex)
    On Error GoTo 0

Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Property

Public Sub Add(ByVal Cll As TableManager.CellClass)

    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Add"
    On Error GoTo ErrorHandler
    
    If Cll.HeaderText <> "" Then pAllCells.Add Cll.Name, Cll
        
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' Add

Public Function Exists(ByVal vIndex) As Boolean
    Exists = pAllCells.Exists(vIndex)
End Function ' Exists

Public Function ListCells() As Variant: Set ListCells = pAllCells.Keys: End Function

Public Sub Remove(ByVal vIndex As Variant)
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllCells = Nothing
        Set pAllCells = New Collection
    Else
        If Not pAllCells.Exists(vIndex) Then Err.Raise 9
        pAllCells.Remove vIndex
    End If

Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Remove
