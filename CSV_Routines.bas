Attribute VB_Name = "CSV_Routines"
Option Explicit

Private Const Module_Name As String = "CSV_Routines."

Private Sub InputTable( _
    ByVal FullFileName As String, _
    ByVal TableName As String)
    
    Const RoutineName As String = Module_Name & "InputTable"
    On Error GoTo ErrorHandler
    
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
    Dim Tbl As TableManager.TableClass
    Set Tbl = TableManager.Table(TableName, Module_Name)
    Dim HeaderRng As Range
    Dim NumTableColumns As Long
    Set HeaderRng = Tbl.Table.HeaderRowRange
    NumTableColumns = HeaderRng.Count
    
    Dim NumFileColumns As Long
    NumFileColumns = UBound(Ary, 2)
    
    If NumTableColumns <> NumFileColumns Then
        MsgBox _
            "There are " & NumTableColumns & _
            " columns in the table and " & NumFileColumns & " columns in the input file", _
            vbOKOnly Or vbCritical, _
            "Input File Size Does Not Match"
        Exit Sub
    End If
    
    ' Check that names of the column headers match else exit
    Dim I As Long
    For I = 1 To NumFileColumns
        If HeaderRng(I) <> Ary(1, I) Then
            MsgBox _
                "Column " & I & " is called " & HeaderRng(I) & " in the table and called " & _
                Ary(1, I) & " in the file", _
                vbOKOnly Or vbCritical, _
                "Column Names Do Not Match"
            Exit Sub
        End If
    Next I
    
    ' Delete the table contents but don't delete the entire table
    ClearTable Tbl.Table
    
    ' copy the new contents
    TableManager.CopyToTable TableName, Ary
    
    ' reset the table dimensions
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub

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

Sub OutputTable( _
    ByVal FullFileName As String, _
    ByVal TableName As String)
    
    Const RoutineName As String = Module_Name & "OutputTable"
    On Error GoTo ErrorHandler
    
    
    Dim Tbl As TableManager.TableClass
    Set Tbl = TableManager.Table(TableName, Module_Name)
    
    Dim Sht As Worksheet
    Set Sht = Tbl.Worksheet
    
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

Public Sub HandleCopyFetch(ByVal TableName As String)
    Const RoutineName As String = Module_Name & "HandleCopyFetch"
    On Error GoTo ErrorHandler
    
    Dim Tbl As TableManager.TableClass
    Set Tbl = TableManager.Table(TableName, Module_Name)
    
    Dim View As I_CopyFetchFormView
    Dim LogoImage As control
    Dim Top As Single
    Dim LogoHeight As Single
    Dim LogoWidth As Single
    Const StandardGap As Long = 12 ' TODO StandardGap is copied from FormClass; need to set it one place
    
    If LogoFileExists Then
        Set View = EnhancedDataBaseForm
        Set LogoImage = Logo(View.Form)
        LogoWidth = LogoImage.Width
        LogoHeight = LogoImage.Height
        LogoImage.Top = StandardGap
        Top = LogoHeight + 2 * StandardGap
        ' TODO Set up the button colors and color change with mouseover
        ' See FormClass.BuildOneButton
    Else
        Set View = New PlainDataBaseForm
        LogoWidth = 0
        LogoHeight = 0
    End If
    
    ' Add the texture
    Texture View.Form
    
    With New CopyFetchClass
    
        Set View.Model = .Self
        
        Dim FullFileName As String
        FullFileName = GetFullFileName(TableName)
        
        View.Destination = FullFileName
        View.Source = FullFileName

        If View.ShowDialog(.Self) Then
            If .CopyClicked Then
                OutputTable FullFileName, TableName
            End If
            
            If .FetchClicked Then
                InputTable FullFileName, TableName
            End If
            
            If .ChangeFileClicked Then
            End If
            
        End If
    
    End With 'model goes out of scope

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub


