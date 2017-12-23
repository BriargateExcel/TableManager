Attribute VB_Name = "Module1"
Option Explicit

Private Const Module_Name As String = "Module1."

Public Sub Auto_Open()
    TableManager.AutoOpen ThisWorkbook
End Sub

Public Sub BuildDataDescriptionTable()
    
    If Not TableManager.TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    TableManager.BuildParameterTableOnWorksheet TableManager.mainworkbook
End Sub


Public Sub ExtendDataValidation()

    If Not TableManager.TableDataCollected Then
        MsgBox "Build the tables first"
        Exit Sub
    End If
    
    TableManager.ExtendDataValidationThroughAllTables TableManager.mainworkbook
    
End Sub
