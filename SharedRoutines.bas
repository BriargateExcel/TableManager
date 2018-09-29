Attribute VB_Name = "SharedRoutines"
'@Folder("TableManager.Main")

Option Explicit

Private Const Module_Name As String = "SharedRoutines."

Public Function ActiveCellTableName() As String
    '   Function returns table name if active cell is in a table and
    '   vbnullstring if it isn't.

    ActiveCellTableName = vbNullString

    '   Statement produces error when active cell is not in a table.
    On Error Resume Next
    ActiveCellTableName = ActiveCellListObject.Name

    On Error GoTo 0                              ' Reset the error handling
End Function                                     ' ActiveCellTableName

Public Function ActiveCellWorksheet() As Worksheet
    '   Function returns worksheet of active cell and Nothing if it isn't.

    Set ActiveCellWorksheet = Nothing

    '   Statement produces error when active cell is not in a table.
    On Error Resume Next
    Set ActiveCellWorksheet = ActiveCell.Worksheet

    On Error GoTo 0                              ' Reset the error handling
End Function                                     ' ActiveCellWorksheet

Public Function ActiveCellListObject() As ListObject
    '   Function returns the ListObject if active cell is in a table and
    '   Nothing if it isn't.

    Set ActiveCellListObject = Nothing

    '   Statement produces error when active cell is not in a table.
    On Error Resume Next
    Set ActiveCellListObject = ActiveCell.ListObject

    On Error GoTo 0                              ' Reset the error handling
End Function                                     ' ActiveCellListobject

Public Function CheckForVBAProjectAccessEnabled(ByVal WkBk As Workbook) As Boolean

    Const RoutineName As String = Module_Name & "CheckForVBAProjectAccessEnabled"
    On Error GoTo ErrorHandler
    
    On Error GoTo ErrorHandler
    
    Dim VBP As VBProject
    
    If Val(Application.Version) >= 10 Then
        Set VBP = WkBk.VBProject
        CheckForVBAProjectAccessEnabled = True
    Else
        MsgBox "This application must be run on Excel 2002 or greater", _
               vbCritical, "Excel Version Check"
        GoTo ErrorHandler
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' CheckForVBAProjectAccessEnabled

Public Function InScope( _
       ByVal ModuleList As Variant, _
       ByVal ModuleName As String _
       ) As Boolean

    '   Uses the name of the module where InScope is called
    '   Filters the name against the list of valid module names
    '   Returns true if the Filter result has any entries
    '   In other words, returns True if ModuleName is found in ModuleList
    
    Const RoutineName As String = "InScope"
    On Error GoTo ErrorHandler
    
    '     Log RoutineName & ":    " & ModuleName

    Dim OneDimArray() As Variant
    
    Const ThisRoutine As String = Module_Name & "InScope"
    On Error GoTo ErrorHandler
    
    Dim NumDim As Long: NumDim = NumberOfArrayDimensions(ModuleList)
    
    If NumDim > 2 Then
        MsgBox "InScope cannot handle arrays with " & _
               "more than 2 dimensions", _
               vbOKOnly Or vbCritical, _
               "NumDim Error"
        Exit Function
    End If
    
    If NumDim = 2 Then
        Dim I As Long
        ReDim OneDimArray(UBound(ModuleList, 1) - 1) As Variant
        For I = 0 To UBound(ModuleList, 1) - 1
            OneDimArray(I) = ModuleList(I + 1, 1)
        Next I
        
        InScope = _
                (UBound( _
                 Filter(OneDimArray, _
                        ModuleName, _
                        True, _
                        CompareMethod.BinaryCompare) _
                 ) > -1)
        Exit Function
    End If

    InScope = _
            (UBound( _
             Filter(ModuleList, _
                    ModuleName, _
                    True, _
                    CompareMethod.BinaryCompare) _
             ) > -1)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, ThisRoutine & "." & RoutineName, Err.Description

End Function                                     ' InScope

Public Sub ShowAnyForm( _
       ByVal FormName As String, _
       Optional ByVal Modal As FormShowConstants = vbModal)
    
    ' http://www.cpearson.com/Excel/showanyform.htm

    Const RoutineName As String = Module_Name & "ShowAnyForm"
    On Error GoTo ErrorHandler
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShowAnyForm
    ' This procedure will show the UserForm named in FormName, either modally or
    ' modelessly, as indicated by the value of Modal.  If a form is already loaded,
    ' it is reshown without unloading/reloading the form.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Obj As Object
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through the VBA.UserForm object (works like
    ' a collection), to see if the form named by
    ' FormName is already loaded. If so, just call
    ' Show and exit the procedure. If it is not loaded,
    ' add it to the VBA.UserForms object and then
    ' show it.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Obj In VBA.UserForms
        If StrComp(Obj.Name, FormName, vbTextCompare) = 0 Then
            '           ''''''''''''''''''''''''''''''''''''
            '           ' START DEBUGGING/ILLUSTRATION ONLY
            '           ''''''''''''''''''''''''''''''''''''
            '           Obj.Label1.Caption = "Form Already Loaded"
            '           ''''''''''''''''''''''''''''''''''''
            '           ' END DEBUGGING/ILLUSTRATION ONLY
            '           ''''''''''''''''''''''''''''''''''''
            Obj.Show Modal
            Exit Sub
        End If
    Next Obj

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If we make it here, the form named by FormName was
    ' not loaded, and thus not found in VBA.UserForms.
    ' Call the Add method of VBA.UserForms to load the
    ' form and then call Show to show the form.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With VBA.UserForms
        On Error Resume Next
        Err.Clear
        Set Obj = .Add(FormName)
        If Err.Number <> 0 Then
            MsgBox "Err: " & CStr(Err.Number) & "   " & Err.Description
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''
        ' START DEBUGGING/ILLUSTRATION ONLY
        ''''''''''''''''''''''''''''''''''''
        Obj.Label1.Caption = "Form Loaded By ShowAnyForm"
        ''''''''''''''''''''''''''''''''''''
        ' END DEBUGGING/ILLUSTRATION ONLY
        ''''''''''''''''''''''''''''''''''''
        Obj.Show Modal
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
   
End Sub                                          ' ShowAnyForm

Public Sub RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal Proc As String, _
       ByVal Desc As String)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Reraises an error and adds line number and current procedure name
    
    Dim SourceOfError As String
    
    ' Check if procedure where error occurs the line no and proc details
    If Src = ThisWorkbook.VBProject.Name Then
        ' Add error line number if present
        If Erl <> 0 Then
            SourceOfError = vbCrLf & "Line no: " & Erl & " "
        End If
   
        ' Add procedure to source
        SourceOfError = SourceOfError & vbCrLf & Proc
        
    Else
        ' If error has already been raised then just add on procedure name
        SourceOfError = Src & vbCrLf & Proc
    End If
    
    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise ErrorNo, SourceOfError, Desc
    
End Sub                                          ' RaiseError

Public Sub DisplayError(ByVal Procname As String)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Displays the error when it reaches the topmost sub
    ' Note: You can add a call to logging from this sub

    Dim Msg As String
    Msg = "The following error occurred: " & vbCrLf & Err.Description _
        & vbCrLf & vbCrLf & "Error Location is: "

    Msg = Msg + Err.Source & vbCrLf & Procname   ' & " " & src & " " & desc

    ' Display message
    MsgBox Msg, Title:="Error"
End Sub                                          ' DisplayError

Public Sub Log(ParamArray Msg() As Variant)
    ' http://analystcave.com/vba-proper-vba-error-handling/
    ' https://excelmacromastery.com/vba-error-handling/
    
    Dim Filename As String
    Filename = GetMainWorkbook.Path & "\error_log.txt"
    Dim MsgString As Variant
    Dim I As Long
    
    Exit Sub

    ' Archive file at certain size
    If FileLen(Filename) > 20000 Then
        FileCopy Filename, _
                 Replace(Filename, ".txt", _
                         Format$(Now, "ddmmyyyy hhmmss.txt"))
        Kill Filename
    End If

    ' Open the file to write
    Dim FileNumber As Long
    FileNumber = FreeFile
    Open Filename For Append As #FileNumber

    MsgString = Msg(LBound(Msg))
    For I = LBound(Msg) + 1 To UBound(Msg)
        MsgString = "," & MsgString & Msg(I)
    Next I

    Print #FileNumber, Now & ":  " & MsgString

    Close #FileNumber
    
End Sub                                          ' Log

Public Function TimeFormat(ByVal Dt As Date) As String
    TimeFormat = Format$(Dt, "hh:mm:ss")
End Function                                     ' TimeFormat

Public Function NumberOfArrayDimensions(ByVal Arr As Variant) As Long
    ' http://www.cpearson.com/excel/vbaarrays.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Long
    Dim Res As Long
    
    Const RoutineName As String = Module_Name & "NumberOfArrayDimensions"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Res = UBound(Arr, 2)                         ' If Arr has only one element, this will fail
    If Err.Number <> 0 Then
        NumberOfArrayDimensions = 1
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0
    
    On Error GoTo ErrorHandler
    
    NumberOfArrayDimensions = Ndx - 1

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' NumberOfArrayDimensions

Public Function HasVal(ByVal Target As Range) As Boolean

    Const RoutineName As String = Module_Name & "HasVal"
    On Error GoTo ErrorHandler
    
    Dim Vbl As Variant
    
    On Error Resume Next
    
    Vbl = Target.Validation.Type
    HasVal = (Err.Number = 0)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' HasVal

Public Function ConvertToLetter(ByVal Col As Long) As String
    Dim Alpha As Long
    Dim Remainder As Long

    Alpha = Int(Col / 27)
    Remainder = Col - (Alpha * 26)
    
    If Alpha > 0 Then
        ConvertToLetter = Chr$(Alpha + 64)
    End If
    
    If Remainder > 0 Then
        ConvertToLetter = ConvertToLetter & Chr$(Remainder + 64)
    End If

End Function                                     ' ConvertToLetter

Public Sub ClearTable(ByVal LstObj As ListObject)
    If LstObj.ListRows.Count > 1 Then
        LstObj.DataBodyRange.Delete
    ElseIf LstObj.ListRows.Count > 0 Then
        LstObj.DataBodyRange.Clear
    End If

End Sub                                          ' ClearTable

Public Function Contains( _
       ByVal Coll As Object, _
       ByVal NamePotentiallyInCollection As String _
       ) As Boolean
    
    Dim Obj As Object

    On Error Resume Next
    
    Set Obj = Coll(NamePotentiallyInCollection)
    Contains = (Err.Number = 0)
End Function                                     ' Contains

Public Function FindLastColumnNumber( _
       ByVal RowNumber As Long, _
       ByVal Sht As Worksheet _
       ) As Long
    
    FindLastColumnNumber = Sht.Cells(RowNumber, Sht.Columns.Count).End(xlToLeft).Column
End Function                                     ' FindLastColumnNumber

Public Function FindLastRow( _
       ByVal ColLetter As String, _
       ByVal RowNumber As Long, _
       ByVal Sheet As Worksheet _
       ) As Long
    Dim RegionRow As Long: RegionRow = Sheet.Range(ColLetter & RowNumber).CurrentRegion.Rows.Count
    Dim ColumnRow As Long: ColumnRow = Sheet.Range(ColLetter & Sheet.Rows.Count).End(xlUp).Row
    Dim ColumnNumber As Long: ColumnNumber = Sheet.Range(ColLetter & 1).Column
    Dim I As Long
    Dim CurrentCell As Range

    If RegionRow = ColumnRow Then
        FindLastRow = ColumnRow
    Else
        For I = Application.WorksheetFunction.Max(ColumnRow, RegionRow) To Application.WorksheetFunction.Min(ColumnRow, RegionRow) Step -1
            Set CurrentCell = Sheet.Cells(I, ColumnNumber)
            If Not IsEmpty(CurrentCell) Then
                FindLastRow = I
                Exit For
            End If
        Next I
    End If
End Function                                     ' FindLastRow

Public Function FileExists(ByVal FullFileName As String) As Boolean
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    FileExists = FSO.FileExists(FullFileName)
End Function

Public Sub CenterMe(ByVal Frm As Object)
    With Frm
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .InsideWidth)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .InsideHeight)
    End With
End Sub

Public Function IsArrayAllocated(ByVal Arr As Variant) As Boolean
    ' http://www.cpearson.com/excel/isarrayallocated.aspx
    On Error Resume Next
    IsArrayAllocated = _
                     IsArray(Arr) And _
                     Not IsError(LBound(Arr, 1)) And _
                     LBound(Arr, 1) <= UBound(Arr, 1)
    On Error GoTo 0
End Function

