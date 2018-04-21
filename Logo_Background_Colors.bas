Attribute VB_Name = "Logo_Background_Colors"
Option Explicit

Private Const Module_Name As String = "Logo_Background_Colors."

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.", "FormClass.", "FormRoutines.", "DataBaseRoutines.")
End Function                                     ' ModuleList

Public Sub DisableButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Btn.Enabled = False

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' DisableButton

Public Sub EnableButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' EnableButton

Public Sub HighLightButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Btn.ForeColor = TableManager.DarkestColorValue
    Btn.BackColor = TableManager.LightestColorValue
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' HighLightButton

Public Sub HighLightControl( _
       ByVal Ctl As control, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Ctl.ForeColor = TableManager.DarkestColorValue
    Ctl.BackColor = TableManager.LightestColorValue

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' HighLightControl

Public Sub LowLightButton( _
       ByVal Btn As MSForms.CommandButton, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Btn.ForeColor = TableManager.LightestColorValue
    Btn.BackColor = TableManager.DarkestColorValue
    Btn.Enabled = True

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' LowLightButton

Public Sub LowLightControl( _
       ByVal Ctl As control, _
       ByVal ModuleName As String)
    
    If Ctl Is Nothing Then Exit Sub
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    On Error GoTo ErrorHandler
    
    Debug.Assert TableManager.InScope(ModuleList, ModuleName)
    
    Ctl.ForeColor = TableManager.LightestColorValue
    Ctl.BackColor = TableManager.DarkestColorValue

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' LowLightControl


