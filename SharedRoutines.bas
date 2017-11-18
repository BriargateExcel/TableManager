Attribute VB_Name = "SharedRoutines"
Option Explicit

Private Const Module_Name = "SharedRoutines."

Function ActiveCellTableName() As String
'   Function returns table name if active cell is in a table and
'   "" if it isn't.

    ActiveCellTableName = ""

'   Statement produces error when active cell is not in a table.
    On Error Resume Next
    ActiveCellTableName = ActiveCell.ListObject.Name

    On Error GoTo 0 ' Reset the error handling
End Function ' ActiveCellTableName

Public Function CheckForVBAProjectAccessEnabled(ByVal WkBkName As String) As Boolean

    Dim VBP As Object ' as VBProject
    Dim WkBk As Workbook

    Const RoutineName = Module_Name & "CheckForVBAProjectAccessEnabled"
    On Error GoTo ErrorHandler
    
    Set WkBk = Workbooks(WkBkName)

    If Val(Application.Version) >= 10 Then
        Set VBP = WkBk.VBProject
    Else
        MsgBox "This application must be run on Excel 2002 or greater", _
            vbCritical, "Excel Version Check"
        GoTo ErrorHandler
    End If

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
    Const ThisRoutine = Module_Name & "InScope"
    On Error GoTo ErrorHandler
    
    Dim NumDim As Integer: NumDim = NumberOfArrayDimensions(ModuleList)
    
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

Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, ThisRoutine & "." & RoutineName, Err.Description

End Function ' InScope

Public Function VBAMatch( _
    ByVal Target As Variant, _
    ByVal SearchRange As Range, _
    Optional ByVal TreatAsString As Boolean = False _
    ) As Long

    On Error GoTo NotFound

    If IsDate(Target) And Not TreatAsString Then
        VBAMatch = Application.Match(CLng(Target), SearchRange, 0)
        Exit Function
    Else
        VBAMatch = Application.WorksheetFunction.Match(Target, SearchRange, 0)
        Exit Function
    End If

NotFound:
    VBAMatch = 0

End Function ' VBAMatch

Sub ShowAnyForm(FormName As String, Optional Modal As FormShowConstants = vbModal)
' http://www.cpearson.com/Excel/showanyform.htm

    Const RoutineName = Module_Name & "ShowAnyForm"
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
    
Done:
    Exit Sub
ErrorHandler:
   RaiseError Err.Number, Err.Source, RoutineName, Err.Description
   
End Sub ' ShowAnyForm

Sub RaiseError( _
    ByVal errorno As Long, _
    ByVal src As String, _
    ByVal proc As String, _
    ByVal desc As String)

' https://excelmacromastery.com/vba-error-handling/
' Reraises an error and adds line number and current procedure name
    
    Dim sSource As String
    
    ' Check if procedure where error occurs the line no and proc details
    If src = ThisWorkbook.VBProject.Name Then
        ' Add error line number if present
        If Erl <> 0 Then
            sSource = vbCrLf & "Line no: " & Erl & " "
        End If
   
        ' Add procedure to source
        sSource = sSource & vbCrLf & proc
        
    Else
        ' If error has already been raised then just add on procedure name
        sSource = src & vbCrLf & proc
    End If
    
    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise errorno, sSource, desc
    
End Sub ' RaiseError

Sub DisplayError(ByVal sProcname As String)

' https://excelmacromastery.com/vba-error-handling/
' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub

    Dim sMsg As String
    sMsg = "The following error occurred: " & vbCrLf & Err.Description _
                    & vbCrLf & vbCrLf & "Error Location is: "

    sMsg = sMsg + Err.Source & vbCrLf & sProcname ' & " " & src & " " & desc

    ' Display message
    MsgBox sMsg, Title:="Error"
End Sub ' DisplayError

Sub Log(ParamArray Msg() As Variant)
' http://analystcave.com/vba-proper-vba-error-handling/
' https://excelmacromastery.com/vba-error-handling/
    
    Dim sFilename As String
    sFilename = MainWorkbook.Path & "\error_log.txt"
    Dim MsgString As Variant
    Dim I As Long
    
    Exit Sub

    ' Archive file at certain size
    If FileLen(sFilename) > 20000 Then
        FileCopy sFilename, _
            Replace(sFilename, ".txt", _
                Format(Now, "ddmmyyyy hhmmss.txt"))
        Kill sFilename
    End If

    ' Open the file to write
    Dim filenumber As Integer
    filenumber = FreeFile
    Open sFilename For Append As #filenumber

    MsgString = Msg(LBound(Msg))
    For I = LBound(Msg) + 1 To UBound(Msg)
        MsgString = "," & MsgString & Msg(I)
    Next I

    Print #filenumber, Now & ":  " & MsgString

    Close #filenumber
    
End Sub ' Log

Public Function DateFormat(ByVal Dt As Date) As String
    If Dt = 0 Then
        DateFormat = Empty
    Else
        DateFormat = Format(Dt, "mm/dd/yyyy")
    End If
End Function ' DateFormat

Public Function TimeFormat(ByVal Dt As Date) As String
    TimeFormat = Format(Dt, "hh:mm:ss")
End Function ' TimeFormat

Public Function NumberOfArrayDimensions(Arr As Variant) As Integer
' http://www.cpearson.com/excel/vbaarrays.htm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim Res As Integer
    
    Const RoutineName = Module_Name & "NumberOfArrayDimensions"
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

Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function ' NumberOfArrayDimensions

