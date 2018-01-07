Attribute VB_Name = "Module1"
Option Explicit

Private Const Module_Name As String = "Module1."

Public Sub Auto_Open()
    TableManager.AutoOpen ThisWorkbook
End Sub

Public Sub BuildDataDescriptionTable()
    
    Const RoutineName As String = Module_Name & "BuildDataDescriptionTable"
    On Error GoTo ErrorHandler
    
    If Not TableManager.TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    TableManager.BuildParameterTableOnWorksheet TableManager.mainworkbook
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub


Public Sub ExtendDataValidation()

    Const RoutineName As String = Module_Name & "ExtendDataValidation"
    On Error GoTo ErrorHandler
    
    If Not TableManager.TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    TableManager.ExtendDataValidationThroughAllTables TableManager.mainworkbook
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub
