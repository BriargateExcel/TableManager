Attribute VB_Name = "Main"
'@Folder("TableManager.Main")

Option Explicit

Private Const Module_Name As String = "Module1."

Public Sub Auto_Open()
    SetUpWorkbook ThisWorkbook
End Sub

Public Sub test()
Dim Wkbk As Workbook

    Workbooks.Open "C:\Users\Owner\Documents\Excel\Headcount\Calendars\LM.xlsx"

    Set Wkbk = Workbooks("LM.xlsx")
    SetUpWorkbook Wkbk
End Sub

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
