Attribute VB_Name = "Main"
'@Folder("TableManager.Main")

Option Explicit

Private Const Module_Name As String = "Module1."

Public Sub BuildDataDescriptionTable(ByVal Wkbk As Workbook)
    
    Const RoutineName As String = Module_Name & "BuildDataDescriptionTable"
    On Error GoTo ErrorHandler
    
    If Not TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    BuildParameterTableOnWorksheet Wkbk
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

Public Sub ExtendDataValidation(ByVal Wkbk As Workbook)

    Const RoutineName As String = Module_Name & "ExtendDataValidation"
    On Error GoTo ErrorHandler
    
    If Not TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    ExtendDataValidationThroughAllTables Wkbk
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

