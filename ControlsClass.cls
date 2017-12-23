VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name As String = "ControlsClass."

Private pAllControls As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllControls = New Scripting.Dictionary:
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("FormClass.")
End Function                                     ' ModuleList

Public Property Get Count() As Long: Count = pAllControls.Count: End Property

'@DefaultMember
Public Property Get Item( _
    ByVal vIndex As Variant, _
    ByVal Modulename As String _
    ) As Control
    'Attribute Item.VB_UserMemId = 0

    Const RoutineName As String = Module_Name & "Get Item"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    On Error Resume Next
    Set Item = pAllControls.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllControls(vIndex)
    On Error GoTo 0

    '@Ignore LineLabelNotUsed
Done:
    Exit Property
    '@Ignore LineLabelNotUsed
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Property Get ListControls(): ListControls = pAllControls.Keys: End Property

Public Sub Add( _
    ByVal Ctl As Control, _
    ByVal Modulename As String)
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Add"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    On Error GoTo ErrorHandler
    
    If Ctl.Name <> vbNullString Then pAllControls.Add Ctl.Name, Ctl
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Add

Public Sub Remove( _
    ByVal vIndex As Variant, _
    ByVal Modulename As String)
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Remove"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllControls = Nothing
        Set pAllControls = New Collection
    Else
        If Not pAllControls.Exists(vIndex) Then Err.Raise 9
        pAllControls.Remove vIndex
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' Remove

