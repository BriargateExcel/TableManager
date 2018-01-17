Attribute VB_Name = "CSV_Routines"
Option Explicit

Private Const Module_Name As String = "CSV_Routines."

Sub TestInput()
    Dim Ary() As Variant

    Dim FSO As Scripting.FileSystemObject
    Dim Folder As String
    Dim FullFileName As String
    
    Const RoutineName As String = Module_Name & "TestInput"
    On Error GoTo ErrorHandler
    
    SetMainWorkbook Workbooks("New Hire or Replace List 2017-12-23.xlsm")
    
    Set FSO = New Scripting.FileSystemObject
    Folder = GetMainWorkbook.Path
    FullFileName = FSO.BuildPath(Folder, "CtlAcct")
    
    'check extension and correct if needed
    If InStr(FullFileName, ".csv") = 0 Then
        FullFileName = FullFileName & ".csv"
    Else
        While (Len(FullFileName) - InStr(FullFileName, ".csv")) > 3
            FullFileName = Left(FullFileName, Len(FullFileName) - 1)
        Wend
    End If

    If Not FSO.FileExists(FullFileName) Then
        MsgBox FullFileName & " does not exist", vbOKOnly Or vbCritical, "File Does Not Exist"
        Exit Sub
    End If

    Ary = ArrayFromCSVfile(FullFileName)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

Public Function ArrayFromCSVfile( _
       ByVal FullFileName As String, _
       Optional ByVal RowDelimiter As String = vbCr, _
       Optional ByVal FieldDelimiter = ",", _
       Optional ByVal RemoveQuotes As Boolean = True _
       ) As Variant
    ' https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet
    ' Assumes file name ends with ".csv"
    ' Load a file created by FileToArray into a 2-dimensional array
    ' The file name is specified by strName, and it is exected to exist
    ' in the user's temporary folder. This is a deliberate restriction:
    ' it's always faster to copy remote files to a local drive than to
    ' edit them across the network

    ' RemoveQuotes=TRUE strips out the double-quote marks (Char 34) that
    ' encapsulate strings in most csv files.

    Const RoutineName As String = Module_Name & "ArrayFromCSVfile"
    On Error GoTo ErrorHandler
    
    Dim FSO As Scripting.FileSystemObject
    Dim arrData As Variant
    Dim Folder As String

    Set FSO = New Scripting.FileSystemObject
    
    If Not FSO.FileExists(FullFileName) Then     ' raise an error?
        Exit Function
    End If

    Application.StatusBar = "Reading the file... (" & FullFileName & ")"

    If Not RemoveQuotes Then
        arrData = Join2d(FSO.OpenTextFile(FullFileName, ForReading).ReadAll, RowDelimiter, FieldDelimiter)
        Application.StatusBar = "Reading the file... Done"
    Else
        ' we have to do some allocation here...
        Dim OneLine As String

        OneLine = FSO.OpenTextFile(FullFileName, ForReading).ReadAll
        Application.StatusBar = "Reading the file... Done"

        Application.StatusBar = "Parsing the file..."

        OneLine = Replace$(OneLine, Chr(34) & RowDelimiter, RowDelimiter)
        OneLine = Replace$(OneLine, RowDelimiter & Chr(34), RowDelimiter)
        OneLine = Replace$(OneLine, Chr(34) & FieldDelimiter, FieldDelimiter)
        OneLine = Replace$(OneLine, FieldDelimiter & Chr(34), FieldDelimiter)

        If Right$(OneLine, Len(OneLine)) = Chr(34) Then
            OneLine = Left$(OneLine, Len(OneLine) - 1)
        End If

        If Left$(OneLine, 1) = Chr(34) Then
            OneLine = Right$(OneLine, Len(OneLine) - 1)
        End If

        Application.StatusBar = "Parsing the file... Done"
        arrData = Split2d(OneLine, RowDelimiter, FieldDelimiter)
        OneLine = ""
    End If

    Application.StatusBar = False

    Set FSO = Nothing
    ArrayFromCSVfile = arrData
    Erase arrData

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Public Function Split2d(ByVal InputString As String, _
                        Optional ByVal RowDelimiter As String = vbCrLf, _
                        Optional ByVal FieldDelimiter = vbTab, _
                        Optional ByVal CoerceLowerBound As Long = 1 _
                        ) As Variant
    ' https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet
    ' Split up a string into a 2-dimensional array.

    ' Works like VBA.Strings.Split, for a 2-dimensional array.
    ' Check your lower bounds on return: never assume that any array in
    ' VBA is zero-based, even if you've set Option Base 0
    ' If in doubt, coerce the lower bounds to 0 or 1 by setting
    ' CoerceLowerBound
    ' Note that the default delimiters are those inserted into the
    '  string returned by ADODB.Recordset.GetString

    Const RoutineName As String = Module_Name & "Split2d"
    On Error GoTo ErrorHandler
    
    ' Coding note: we're not doing any string-handling in VBA.Strings -
    ' allocating, deallocating and (especially!) concatenating are SLOW.
    ' We're using the VBA Join & Split functions ONLY. The VBA Join,
    ' Split, & Replace functions are linked directly to fast (by VBA
    ' standards) functions in the native Windows code. Feel free to
    ' optimise further by declaring and using the Kernel string functions
    ' if you want to.

    ' ** THIS CODE IS IN THE PUBLIC DOMAIN **
    '    Nigel Heffernan   Excellerando.Blogspot.com

    Dim I   As Long
    Dim J   As Long

    Dim i_n As Long
    Dim j_n As Long

    Dim FirstRow As Long
    Dim LastRow As Long
    Dim FirstColumn As Long
    Dim LastColumn As Long

    Dim ArrayOfRows As Variant
    Dim OneRow As Variant

    ArrayOfRows = Split(InputString, RowDelimiter)

    FirstRow = LBound(ArrayOfRows)
    LastRow = UBound(ArrayOfRows)

    If Len(VBA.LenB(ArrayOfRows(LastRow))) <= 1 Then
        ' clip out empty last row: a common artifact in data
        'loaded from files with a terminating row delimiter
        LastRow = LastRow - 1
    End If

    I = FirstRow
    OneRow = Split(ArrayOfRows(I), FieldDelimiter)

    FirstColumn = LBound(OneRow)
    LastColumn = UBound(OneRow)

    If VBA.LenB(OneRow(LastColumn)) <= 0 Then
        ' ! potential error: first row with an empty last field...
        LastColumn = LastColumn - 1
    End If

    i_n = CoerceLowerBound - FirstRow
    j_n = CoerceLowerBound - FirstColumn

    ReDim arrData(FirstRow + i_n To LastRow + i_n, FirstColumn + j_n To LastColumn + j_n)

    ' As we've got the first row already... populate it
    ' here, and start the main loop from lbound+1

    For J = FirstColumn To LastColumn
        arrData(FirstRow + i_n, J + j_n) = OneRow(J)
    Next J

    For I = FirstRow + 1 + i_n To LastRow + i_n Step 1

        OneRow = Split(ArrayOfRows(I), FieldDelimiter)

        For J = FirstColumn To LastColumn Step 1
            arrData(I + i_n, J + j_n) = OneRow(J)
        Next J

        Erase OneRow

    Next I

    Erase ArrayOfRows

    Application.StatusBar = False

    Split2d = arrData

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Public Function Join2d(ByRef InputArray As Variant, _
                       Optional RowDelimiter As String = vbCr, _
                       Optional FieldDelimiter = vbTab, _
                       Optional SkipBlankRows As Boolean = False _
                       ) As String
    ' https://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet
    ' Join up a 2-dimensional array into a string. Works like the standard
    '  VBA.Strings.Join, for a 2-dimensional array.
    ' Note that the default delimiters are those inserted into the string
    '  returned by ADODB.Recordset.GetString

    Const RoutineName As String = Module_Name & "Split2d"
    On Error GoTo ErrorHandler
    
    ' Coding note: we're not doing any string-handling in VBA.Strings -
    ' allocating, deallocating and (especially!) concatenating are SLOW.
    ' We're using the VBA Join & Split functions ONLY. The VBA Join,
    ' Split, & Replace functions are linked directly to fast (by VBA
    ' standards) functions in the native Windows code. Feel free to
    ' optimise further by declaring and using the Kernel string functions
    ' if you want to.

    ' ** THIS CODE IS IN THE PUBLIC DOMAIN **
    '   Nigel Heffernan   Excellerando.Blogspot.com

    Dim I As Long
    Dim J As Long

    Dim i_lBound As Long
    Dim i_uBound As Long
    Dim j_lBound As Long
    Dim j_uBound As Long

    Dim arrTemp1() As String
    Dim arrTemp2() As String

    Dim strBlankRow As String

    i_lBound = LBound(InputArray, 1)
    i_uBound = UBound(InputArray, 1)

    j_lBound = LBound(InputArray, 2)
    j_uBound = UBound(InputArray, 2)

    ReDim arrTemp1(i_lBound To i_uBound)
    ReDim arrTemp2(j_lBound To j_uBound)

    For I = i_lBound To i_uBound

        For J = j_lBound To j_uBound
            arrTemp2(J) = InputArray(I, J)
        Next J

        arrTemp1(I) = Join(arrTemp2, FieldDelimiter)

    Next I

    If SkipBlankRows Then

        If Len(FieldDelimiter) = 1 Then
            strBlankRow = String(j_uBound - j_lBound, FieldDelimiter)
        Else
            For J = j_lBound To j_uBound
                strBlankRow = strBlankRow & FieldDelimiter
            Next J
        End If

        Join2d = Replace(Join(arrTemp1, RowDelimiter), strBlankRow, RowDelimiter, "")
        I = Len(strBlankRow & RowDelimiter)

        If Left(Join2d, I) = strBlankRow & RowDelimiter Then
            Mid$(Join2d, 1, I) = ""
        End If

    Else

        Join2d = Join(arrTemp1, RowDelimiter)

    End If

    Erase arrTemp1

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function

Sub TestOutput()
    Const RoutineName As String = Module_Name & "TestOutput"
    On Error GoTo ErrorHandler
    
    SetMainWorkbook Workbooks("New Hire or Replace List 2017-12-23.xlsm")
    
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

    SaveAsCSV Ary, "CtlAcct"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub

Public Sub SaveAsCSV( _
       ByRef MyArray() As Variant, _
       ByVal FileName As String, _
       Optional ByVal LowerBound As Long = 1, _
       Optional ByVal Delimiter As String = ",")
    ' https://stackoverflow.com/questions/4191560/create-csv-from-array-in-vba
    ' SaveAsCSV saves an array as csv file. Choosing a delimiter different as a comma, is optional.
    '
    ' Syntax:
    ' SaveAsCSV dMyArray, sMyFileName, [sMyDelimiter]
    '
    ' Examples:
    ' SaveAsCSV dChrom, app.path & "\Demo.csv"       --> comma as delimiter
    ' SaveAsCSV dChrom, app.path & "\Demo.csv", ";"  --> semicolon as delimiter
    '
    ' Rev. 1.00 [8 jan 2003]
    ' written by P. Wester
    ' wester@kpd.nl

    Dim n As Long                                'counter
    Dim M As Long                                'counter
    Dim CSV As String                            'csv string to print

    Const RoutineName As String = Module_Name & "SaveACSV"
    On Error GoTo ErrorHandler
    
    Dim FSO As Scripting.FileSystemObject
    Dim Folder As String
    Dim FullFileName As String

    Set FSO = New Scripting.FileSystemObject
    Folder = GetMainWorkbook.Path
    FullFileName = FSO.BuildPath(Folder, FileName)
    
    'check extension and correct if needed
    If InStr(FullFileName, ".csv") = 0 Then
        FullFileName = FullFileName & ".csv"
    Else
        While (Len(FullFileName) - InStr(FullFileName, ".csv")) > 3
            FullFileName = Left(FullFileName, Len(FullFileName) - 1)
        Wend
    End If

    Dim Response As String
    If FSO.FileExists(FullFileName) Then
        Response = MsgBox(FullFileName & " already exists. Overwrite?", vbYesNo Or vbExclamation, "File Exists")
        If Response = vbNo Then Exit Sub
    End If
    
    Dim UpperBound1 As Long
    Dim UpperBound2 As Long

    If NumberOfArrayDimensions(MyArray()) = 1 Then '1 dimensional
        Open FullFileName For Output As #7
        ' TODO Check the default value of lower bound of one=-dimensional array
        If LowerBound = 1 Then
            UpperBound1 = UBound(MyArray(), 1)
        Else
            UpperBound1 = UBound(MyArray(), 1) - 1
        End If
        For n = LowerBound To UpperBound1
            Print #7, Format(MyArray(n, 0), "0.000000E+00")
        Next n
        Close #7

    ElseIf NumberOfArrayDimensions(MyArray()) = 2 Then '2 dimensional
        Open FullFileName For Output As #7
        
        If LowerBound = 1 Then
            UpperBound1 = UBound(MyArray(), 1)
            UpperBound2 = UBound(MyArray(), 2)
        Else
            UpperBound1 = UBound(MyArray(), 1) - 1
            UpperBound2 = UBound(MyArray(), 2) - 1
        End If
        
        For n = LowerBound To UpperBound1
            CSV = ""
            For M = LowerBound To UpperBound2
                CSV = CSV & Format(MyArray(n, M)) & Delimiter
            Next M
            CSV = Left(CSV, Len(CSV) - 1)        'remove last Delimiter
            Print #7, CSV
        Next n
        Close #7
    Else
        Stop
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    Close #7
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub


