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

Private Const Module_Name As String = "CellsClass."

Private pAllCells As Scripting.Dictionary

' TODO Examine the design of the plural classes - see notes below
' All the plural classes are built on a Dictionary
' The plural classes: Cells, Controls, Events, Tables, Worksheets
' They all have these and only these routines:
'     Count, Item, Add, Remove, and Class_Initialize
' Challenge:
'   Item returns the singular of the class
'   Add takes the singular of the class as input
' Straightforward:
'   Count takes no parameters
'   Remove takes the same parameters in all plural classes
' Is there any advantage in building them as Interface/Implements classes?
' CellsClass      is only used in TableClass
' ControlsClass   is only used in FormClass
' EventsClass     is only used in FormClass
' TablesClass     is only used in TableRoutines
' WorksheetsClass is only used in WorksheetRoutines

Private Sub Class_Initialize()
    Set pAllCells = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("FormClass.", "FormRoutines.", "TableRoutines.", "CellsClass.", "TableClass.", "ParameterRoutines.")
End Function                                     ' ModuleList

Public Property Get Count() As Long: Count = pAllCells.Count: End Property

'TODO Figure out the Rubberduck annotation for a default member
'@DefaultMember
Public Property Get Item( _
    ByVal vIndex As Variant, _
    ByVal Modulename As String _
    ) As TableManager.CellClass
    'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName As String = Module_Name & "Get Item"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, Modulename)
    
    On Error Resume Next
    Set Item = pAllCells.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllCells(vIndex)
    On Error GoTo 0

    Debug.Assert Not Item Is Nothing

    '@Ignore LineLabelNotUsed
Done:
    Exit Property
    '@Ignore LineLabelNotUsed
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property

Public Sub Add( _
    ByVal Cll As TableManager.CellClass, _
    ByVal Modulename As String)

    Debug.Assert TableManager.Initializing

    Const RoutineName As String = Module_Name & "Add"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, Modulename)
    
    On Error GoTo ErrorHandler
    
    If Cll.HeaderText <> vbNullString Then pAllCells.Add Cll.Name, Cll
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' Add

Public Sub Remove( _
    ByVal vIndex As Variant, _
    ByVal Modulename As String)
    
    Debug.Assert TableManager.Initializing

    Const RoutineName As String = Module_Name & "Remove"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, Modulename)
    
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllCells = Nothing
        Set pAllCells = New Collection
    Else
        If Not pAllCells.Exists(vIndex) Then Err.Raise 9
        pAllCells.Remove vIndex
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Remove

Public Function Exists( _
    ByVal vIndex As Variant, _
    ByVal Modulename As String _
    ) As Boolean
    
    Const RoutineName As String = Module_Name & "Exists"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, Modulename)
    
    Exists = pAllCells.Exists(vIndex)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' Exists



