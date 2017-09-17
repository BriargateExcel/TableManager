Attribute VB_Name = "ControlRoutines"
Option Explicit

Private pAllCtls As TableManager.ControlsClass

Public Sub ControlAdd( _
    ByVal Val As TableManager.ControlClass, _
    ByVal ModuleName As String)
    
    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    pAllCtls.Add Val
End Sub

Private Function ControlModuleList() As Variant
    ControlModuleList = Array("ControlsClass.", "FormClass.", "EventHandler.")
End Function

Public Sub ControlSetNewClass(ByVal ModuleName As String)
    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    Set pAllCtls = New TableManager.ControlsClass
End Sub

Public Function ControlItem( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As TableManager.ControlClass

    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    Set ControlItem = pAllCtls(Val)
End Function

Public Function ControlCount(ByVal ModuleName As String) As Long
    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    ControlCount = pAllCtls.Count
End Function

Public Function ControlExists( _
    ByVal Val As Variant, _
    ByVal ModuleName As String _
    ) As Boolean
    
    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    ControlExists = pAllCtls.Exists(Val)
End Function

Public Sub ControlRemove( _
    ByVal Val As Variant, _
    ByVal ModuleName As String)
    
    Debug.Assert TableManager.InScope(ControlModuleList, ModuleName)
    pAllCtls.Remove Val
End Sub

Public Function NewControlClass() As TableManager.ControlClass
    Set NewControlClass = New TableManager.ControlClass
End Function


