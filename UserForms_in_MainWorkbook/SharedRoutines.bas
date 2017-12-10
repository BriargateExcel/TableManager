Attribute VB_Name = "SharedRoutines"
Option Explicit

Private Const Module_Name As String = "SharedRoutines."

Public Function ActiveCellTableName() As String
'   Function returns table name if active cell is in a table and
'   vbnullstring if it isn't.

    ActiveCellTableName = vbNullString

'   Statement produces error when active cell is not in a table.
    On Error Resume Next
    ActiveCellTableName = ActiveCell.ListObject.Name

    On Error GoTo 0 ' Reset the error handling
End Function ' ActiveCellTableName

Public Function CheckForVBAProjectAccessEnabled(ByVal WkBkName As String) As Boolean

    Dim VBP As Object ' as VBProject
    Dim WkBk As Workbook

    Const RoutineName As String = Module_Name & "CheckForVBAProjectAccessEnabled"
    On Error GoTo ErrorHandler
    
    Set WkBk = Workbooks(WkBkName)

    If Val(Application.VERSION) >= 10 Then
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

End Function ' CheckForVBAProjectAccessEnabled

Public Function InScope( _
    ByVal ModuleList As Variant, _
    ByVal ModuleName As String, _
    ByVal RoutineName As String _
    ) As Boolean

'   Uses the name of the module where InScope is called
'   Filters the name against the list of valid module names
'   Returns true if the Filter result has any entries
'   In other words, returns True if ModuleName is found in ModuleList

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

End Function ' InScope

'Public Sub ShowAnyForm( _
'    ByVal FormName As String, _
'    Optional ByVal Modal As FormShowConstants = vbModal)
'' http://www.cpearson.com/Excel/showanyform.htm
'
'    Const RoutineName As String = Module_Name & "ShowAnyForm"
'    On Error GoTo ErrorHandler
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' ShowAnyForm
'    ' This procedure will show the UserForm named in FormName, either modally or
'    ' modelessly, as indicated by the value of Modal.  If a form is already loaded,
'    ' it is reshown without unloading/reloading the form.
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''    Dim Obj As Object
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Loop through the VBA.UserForm object (works like
'    ' a collection), to see if the form named by
'    ' FormName is already loaded. If so, just call
'    ' Show and exit the procedure. If it is not loaded,
'    ' add it to the VBA.UserForms object and then
'    ' show it.
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    Dim I As Long
'    Dim Obj As Object
'    For I = 1 To MainWorkbook.VBProject.VBComponents.Count
'        Set Obj = MainWorkbook.VBProject.VBComponents.Item(I)
'        If StrComp(Obj.Name, FormName, vbTextCompare) = 0 Then
'            Obj.Show Modal
'            Exit Sub
'        End If
'    Next I
''    Dim Tbl As TableManager.TableClass
''    For I = 0 To TableManager.TableCount(Module_Name) - 1
''        Set Tbl = TableManager.TableItem(I, Module_Name)
''        If Tbl.Form.Name = FormName Then
''            Dim Frm As MSForms.UserForm
''            Set Frm = Tbl.Form.FormObj
''            Frm.Show Modal
''        End If
''    Next I
'
'
'
''    For Each Obj In VBA.UserForms
''        If StrComp(Obj.Name, FormName, vbTextCompare) = 0 Then
'''           ''''''''''''''''''''''''''''''''''''
'''           ' START DEBUGGING/ILLUSTRATION ONLY
'''           ''''''''''''''''''''''''''''''''''''
'''           Obj.Label1.Caption = "Form Already Loaded"
'''           ''''''''''''''''''''''''''''''''''''
'''           ' END DEBUGGING/ILLUSTRATION ONLY
'''           ''''''''''''''''''''''''''''''''''''
''            Obj.Show Modal
''            Exit Sub
''        End If
''    Next Obj
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' If we make it here, the form named by FormName was
'    ' not loaded, and thus not found in VBA.UserForms.
'    ' Call the Add method of VBA.UserForms to load the
'    ' form and then call Show to show the form.
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''    With VBA.UserForms
''        On Error Resume Next
''        Err.Clear
''        Set Obj = .Add(FormName)
''        If Err.Number <> 0 Then
''            MsgBox "Err: " & CStr(Err.Number) & "   " & Err.Description
''            Exit Sub
''        End If
''        ''''''''''''''''''''''''''''''''''''
''        ' START DEBUGGING/ILLUSTRATION ONLY
''        ''''''''''''''''''''''''''''''''''''
''        Obj.Label1.Caption = "Form Loaded By ShowAnyForm"
''        ''''''''''''''''''''''''''''''''''''
''        ' END DEBUGGING/ILLUSTRATION ONLY
''        ''''''''''''''''''''''''''''''''''''
''        Obj.Show Modal
''    End With
'
''@Ignore LineLabelNotUsed
'Done:
'    Exit Sub
'ErrorHandler:
'   RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'
'End Sub ' ShowAnyForm

Public Sub RaiseError( _
    ByVal errorno As Long, _
    ByVal src As String, _
    ByVal proc As String, _
    ByVal desc As String)

' https://excelmacromastery.com/vba-error-handling/
' Reraises an error and adds line number and current procedure name
    
    Dim SourceOfError As String
    
    ' Check if procedure where error occurs the line no and proc details
    If src = ThisWorkbook.VBProject.Name Then
        ' Add error line number if present
        If Erl <> 0 Then
            SourceOfError = vbCrLf & "Line no: " & Erl & " "
        End If
   
        ' Add procedure to source
        SourceOfError = SourceOfError & vbCrLf & proc
        
    Else
        ' If error has already been raised then just add on procedure name
        SourceOfError = src & vbCrLf & proc
    End If
    
    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise errorno, SourceOfError, desc
    
End Sub ' RaiseError

Public Sub DisplayError(ByVal Procname As String)

' https://excelmacromastery.com/vba-error-handling/
' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub

    Dim Msg As String
    Msg = "The following error occurred: " & vbCrLf & Err.Description _
                    & vbCrLf & vbCrLf & "Error Location is: "

    Msg = Msg & Err.Source & vbCrLf & Procname ' & " " & src & " " & desc

    ' Display message
    MsgBox Msg, Title:="Error"
End Sub ' DisplayError

Public Sub Log(ParamArray Msg() As Variant)
' http://analystcave.com/vba-proper-vba-error-handling/
' https://excelmacromastery.com/vba-error-handling/
    
    Dim Filename As String
    Filename = MainWorkbook.Path & "\error_log.txt"
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
    Dim filenumber As Long
    filenumber = FreeFile
    Open Filename For Append As #filenumber

    MsgString = Msg(LBound(Msg))
    For I = LBound(Msg) + 1 To UBound(Msg)
        MsgString = "," & MsgString & Msg(I)
    Next I

    Print #filenumber, Now & ":  " & MsgString

    Close #filenumber
    
End Sub ' Log

Public Function TimeFormat(ByVal Dt As Date) As String
    TimeFormat = Format$(Dt, "hh:mm:ss")
End Function ' TimeFormat

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
    Res = UBound(Arr, 2) ' If Arr has only one element, this will fail
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

End Function ' NumberOfArrayDimensions

Public Function HasVal(ByVal Target As Range) As Boolean

    Const RoutineName As String = Module_Name & "HasVal"
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    On Error Resume Next
    
    v = Target.Validation.Type
    If Err.Number = 0 Then
        HasVal = True
    Else
        HasVal = False
    End If

'@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' HasVal










