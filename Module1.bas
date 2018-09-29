Attribute VB_Name = "Module1"
'@Folder("TableManager.Main")

Option Explicit

Private Const Module_Name As String = "Module1."

Public Sub Auto_Open()
    AutoOpen ThisWorkbook
End Sub

Public Sub BuildDataDescriptionTable()
    
    Const RoutineName As String = Module_Name & "BuildDataDescriptionTable"
    On Error GoTo ErrorHandler
    
    If Not TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    BuildParameterTableOnWorksheet GetMainWorkbook
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub


Public Sub ExtendDataValidation()

    Const RoutineName As String = Module_Name & "ExtendDataValidation"
    On Error GoTo ErrorHandler
    
    If Not TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    ExtendDataValidationThroughAllTables GetMainWorkbook
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub
