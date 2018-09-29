Attribute VB_Name = "CSV_Routines"
'@Folder("TableManager.DataBase")

Option Explicit

Private Const Module_Name As String = "CSV_Routines."

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.", "XLAM_Module.", "PlainDataBaseForm.")
End Function                                     ' ModuleList

Public Function GetFullFileName(ByVal Filename As String) As String
    Const RoutineName As String = Module_Name & "GetFullFileName"
    On Error GoTo ErrorHandler
    
    Dim FullFileName As String
    Dim FSO As Scripting.FileSystemObject
    
    Set FSO = New Scripting.FileSystemObject

    FullFileName = FSO.BuildPath(GetWorkBookPath, Filename)
    
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

Public Sub InputTable(ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "InputTable"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Dim FullFileName As String
    FullFileName = GetFullFileName(ActiveCellTableName)
    
    If Not FileExists(FullFileName) Then
        MsgBox FullFileName & " does not exist", vbOKOnly Or vbCritical, "File Does Not Exist"
        Exit Sub
    End If
    Dim Database As I_DataBase
    Set Database = New CSVClass
    Dim Ary As Variant
    With Database
        Ary = .ArrayFromDataBase(FullFileName)
    End With
    
    ' Check that number column headers match else exit
    Dim HeaderRng As Range
    Dim NumTableColumns As Long
    Set HeaderRng = ActiveCellListObject.HeaderRowRange
    NumTableColumns = HeaderRng.Count
    
    Dim NumFileColumns As Long
    NumFileColumns = UBound(Ary, 2)
    
    If NumTableColumns <> NumFileColumns Then
        MsgBox "There are " & _
               NumTableColumns & _
               " columns in the table and " & NumFileColumns & " columns in the input file", _
               vbOKOnly Or vbCritical, _
               "Input File Size Does Not Match"
        Exit Sub
    End If
    
    ' Check that names of the column headers match else exit
    Dim I As Long
    For I = 1 To NumFileColumns
        If HeaderRng(I) <> Ary(1, I) Then
            MsgBox "Column " & I & " is called " & HeaderRng(I) & _
                   " in the table and called " & _
                   Ary(1, I) & " in the file", _
                   vbOKOnly Or vbCritical, _
                   "Column Names Do Not Match"
            Exit Sub
        End If
    Next I
    
    ' Delete the table contents but don't delete the entire table
    ClearTable ActiveCellListObject
    
    ' copy the new contents
    CopyToTable ActiveCellListObject, Ary
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

Public Sub OutputTable(ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "OutputTable"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Dim FullFileName As String
    FullFileName = GetFullFileName(ActiveCellTableName)
    
    Dim Sht As Worksheet
    Set Sht = ActiveCellWorksheet
    
    Dim NumRows As Long
    Dim NumCols As Long
    Dim ColLetter As String
    Dim Rng As String
    NumRows = FindLastRow("A", 1, Sht)
    NumCols = FindLastColumnNumber(1, Sht)
    ColLetter = ConvertToLetter(NumCols)
    Rng = "A1:" & ColLetter & NumRows
    
    Dim Ary() As Variant
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
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

Public Sub ChangeFile( _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "OutputTable"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    MsgBox "Not implemented yet", vbOKOnly, "File Change"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub


