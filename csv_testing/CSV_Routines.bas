Attribute VB_Name = "CSV_Routines"
Option Explicit

Private Const Module_Name As String = "CSV_Routines."

Private Function InputTable(ByVal FullFileName As String) As Variant
    Const RoutineName As String = Module_Name & "InputTable"
    On Error GoTo ErrorHandler
    
    If Not FileExists(FullFileName) Then
        MsgBox FullFileName & " does not exist", vbOKOnly Or vbCritical, "File Does Not Exist"
        Exit Function
    End If
    Dim Database As I_DataBase
    Set Database = New CSVClass
    With Database
        InputTable = .ArrayFromDataBase(FullFileName)
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    DisplayError RoutineName
End Function

Private Function GetFullFileName(ByVal FileName As String) As String
    Const RoutineName As String = Module_Name & "GetFullFileName"
    On Error GoTo ErrorHandler
    
    Dim FullFileName As String
    Dim FSO As Scripting.FileSystemObject
    
    Set FSO = New Scripting.FileSystemObject

    FullFileName = FSO.BuildPath(GetWorkBookPath, FileName)
    
    'check extension and correct if needed
    If InStr(FullFileName, ".csv") = 0 Then
        FullFileName = FullFileName & ".csv"
    Else
        While (Len(FullFileName) - InStr(FullFileName, ".csv")) > 3
            FullFileName = Left$(FullFileName, Len(FullFileName) - 1)
        Wend
    End If

GetFullFileName = FullFileName

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Sub OutputTable(ByVal FullFileName As String)
    Const RoutineName As String = Module_Name & "TestOutput"
    On Error GoTo ErrorHandler
    
    Dim NumRows As Long
    Dim NumCols As Long
    Dim ColLetter As String
    Dim Rng As String
    Dim Sht As Worksheet
    Dim Ary() As Variant
    
    Set Sht = GetMainWorkbook.Worksheets("Control Accounts")
    NumRows = FindLastRow("A", 1, Sht)
    NumCols = FindLastColumnNumber(1, Sht)
    ColLetter = ConvertToLetter(NumCols)
    Rng = "A1:" & ColLetter & NumRows
    Ary = Sht.Range(Rng)

    Dim Response As String
    If FileExists(FullFileName) Then
        Response = MsgBox(FullFileName & " already exists. Overwrite?", vbYesNo Or vbExclamation, "File Exists")
        If Response = vbNo Then Exit Sub
    End If
    
    Dim Database As I_DataBase
    Set Database = New CSVClass
    With Database
        .ArrayToDataBase Ary, FullFileName
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

Public Sub HandleCopyFetch(ByVal TableName As String)
        Dim View As I_CopyFetchFormView
        Set View = New DataBaseForm
        
        With New CopyFetchClass
        
            Set View.Model = .Self
            
            Dim FullFileName As String
            FullFileName = GetFullFileName(TableName)
            
            View.Destination = FullFileName
            View.Source = FullFileName
    
            If View.ShowDialog(.Self) Then
                If .CopyClicked Then
                    OutputTable FullFileName
                End If
                
                If .OtherDestinationClicked Then
                End If
                
                If .FetchClicked Then
                    Dim Ary() As Variant
                    Ary = InputTable(FullFileName)
                    TableManager.CopyToTable TableName, Ary
                End If
                
                If .OtherSourceClicked Then
                End If
                
            End If
        
        End With 'model goes out of scope

End Sub


