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

Private Const Module_Name = "ControlsClass."

Private pAllControls As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllControls = New Scripting.Dictionary:
End Sub ' Class_Initialize

Public Property Get Count() As Integer: Count = pAllControls.Count: End Property

Public Property Get Item(ByVal vIndex As Variant) As Control
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0

    Const RoutineName = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllControls.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    Set Item = pAllControls(vIndex)
    On Error GoTo 0

Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Property Get ListControls(): ListControls = pAllControls.Keys: End Property

Public Sub Add(ByVal Ctl As Control)
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Add"
    On Error GoTo ErrorHandler
    
    If Ctl.Name <> "" Then pAllControls.Add Ctl.Name, Ctl
        
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Add

Public Function Exists(ByVal vIndex) As Boolean
    pAllControls.Exists (vIndex)
End Function ' Exists

Public Sub Remove(ByVal vIndex As Variant)
    
    Debug.Assert Initializing

    Const RoutineName = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllControls = Nothing
        Set pAllControls = New Collection
    Else
        If Not pAllControls.Exists(vIndex) Then Err.Raise 9
        pAllControls.Remove vIndex
    End If

Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' Remove
