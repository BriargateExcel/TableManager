Attribute VB_Name = "Logo_Background_Colors"
Option Explicit

Private Const Module_Name As String = "Logo_Background_Colors."

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.", "FormClass.")
End Function                                     ' ModuleList

Public Sub DisableButton( _
    ByVal Btn As MSForms.CommandButton, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Btn.Enabled = False
End Sub                                          ' DisableButton

Public Sub EnableButton( _
    ByVal Btn As MSForms.CommandButton, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Btn.Enabled = True
End Sub                                          ' EnableButton

Public Sub HighLightButton( _
    ByVal Btn As MSForms.CommandButton, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Btn.ForeColor = TableManager.DarkestColorValue
    Btn.BackColor = TableManager.LightestColorValue
    Btn.Enabled = True
End Sub                                          ' HighLightButton

Public Sub HighLightControl( _
    ByVal Ctl As Control, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Ctl.ForeColor = TableManager.DarkestColorValue
    Ctl.BackColor = TableManager.LightestColorValue
End Sub                                          ' HighLightControl

Public Sub LowLightButton( _
    ByVal Btn As MSForms.CommandButton, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Btn.ForeColor = TableManager.LightestColorValue
    Btn.BackColor = TableManager.DarkestColorValue
    Btn.Enabled = True
End Sub                                          ' LowLightButton

Public Sub LowLightControl( _
    ByVal Ctl As Control, _
    ByVal Modulename As String)
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    Ctl.ForeColor = TableManager.LightestColorValue
    Ctl.BackColor = TableManager.DarkestColorValue
End Sub                                          ' LowLightControl

