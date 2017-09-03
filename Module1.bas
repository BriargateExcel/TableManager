Attribute VB_Name = "Module1"
Option Explicit

Private Const Module_Name = "Module1."

Public Sub Auto_Open()
    TableManager.autoopen (ThisWorkbook.Name)
End Sub

