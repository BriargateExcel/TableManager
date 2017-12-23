Attribute VB_Name = "FormRoutines"
Option Explicit

Private Const Module_Name As String = "FormRoutines."

Private Function ModuleList() As Variant
    ModuleList = Array("EventClass.")
End Function                                     ' ModuleList

Public Function ValidateForm( _
       ByVal Tbl As TableManager.TableClass, _
       ByVal Modulename As String _
       ) As Boolean
    
    Dim Field As TableManager.CellClass
    Dim Intermediate As Boolean: Intermediate = True
    Dim Check As Boolean
    Dim I As Long
    
    Const RoutineName As String = Module_Name & "ValidateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)
    
    On Error GoTo ErrorHandler
    
    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
        
        Select Case Field.CellType

        Case XlDVType.xlValidateInputOnly        ' Validate only when user changes value
            ' Input can be anything; no validation checking possible
            Check = True
            
        Case XlDVType.xlValidateWholeNumber      ' Whole numeric values
            Check = ValidateWholeNumber(Tbl, Field)

        Case XlDVType.xlValidateDecimal          ' Numeric values
            Check = ValidateDecimal(Tbl, Field)

        Case XlDVType.xlValidateList             ' Value must be present in a specified list
            Check = ValidateList(Field)

        Case XlDVType.xlValidateDate             ' Date Values
            Check = ValidateDate(Tbl, Field)

        Case XlDVType.xlValidateTime             ' Time values
            Check = ValidateTime(Tbl, Field)

        Case XlDVType.xlValidateTextLength       ' Length of text
            Check = ValidateTextLength(Tbl, Field)

        Case XlDVType.xlValidateCustom           ' Validate by arbitrary formula
            Check = ValidateCustom(Tbl, Field)

        End Select
        
        If Intermediate Then
            Intermediate = Check
        End If
        
    Next I
    
    ValidateForm = Intermediate
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateForm

Private Function ValString(ByVal Val As Variant) As String
    Const RoutineName As String = Module_Name & "ValString"
    On Error GoTo ErrorHandler
    
    Select Case VarType(Val)
    Case vbEmpty:           ValString = vbNullString
    Case vbNull:            ValString = vbNull
    Case vbInteger:         ValString = Format$(CInt(Val), "0")
    Case vbLong:            ValString = Format$(CLng(Val), "0")
    Case vbSingle:          ValString = Format$(CSng(Val), "0.0")
    Case vbDouble:          ValString = Format$(CDbl(Val), "0.0")
    Case vbString:          ValString = Val
    Case vbObject:          ValString = vbError
    Case vbError:           ValString = vbError
    Case vbBoolean:         ValString = Val
    Case vbVariant:         ValString = vbError
    Case vbDataObject:      ValString = vbError
    Case vbDecimal:         ValString = Format$(CSng(Val), "0.0")
    Case vbByte:            ValString = Val
    Case vbUserDefinedType: ValString = Val
    Case vbArray:           ValString = vbError
    Case vbDate
        If Val = 0 Then
            ValString = vbNullString
        Else
            ValString = Format$(CDate(Val), "mm/dd/yyyy")
        End If
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValString

Private Function CheckRanges( _
        ByVal Val1 As Variant, _
        ByVal Val2 As Variant, _
        ByVal FormVal As Variant, _
        ByVal TableVal As Variant, _
        ByVal Field As TableManager.CellClass _
        ) As Boolean
    
    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "CheckRanges"
    On Error GoTo ErrorHandler
    
    Select Case Field.Operator

    Case XlFormatConditionOperator.xlBetween
        If FormVal >= Val1 And FormVal <= Val2 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must be between " & _
                   Val1 & " and " & Val2, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If
        
    Case XlFormatConditionOperator.xlNotBetween
        If Not (FormVal >= Val1 And FormVal <= Val2) Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must not be between " & _
                   Val1 & " and " & Val2, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlEqual
        If FormVal = Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must equal " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlNotEqual
        If FormVal <> Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must not equal " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlGreater
        If FormVal > Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must be greater than " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlLess
        If FormVal > Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must be greater than " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlGreaterEqual
        If FormVal >= Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must be greater than or equal to " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    Case XlFormatConditionOperator.xlLessEqual
        If FormVal <= Val1 Then
            CheckRanges = True
        Else
            MsgBox Field.Name & " must be less than or equal to " & Val1, _
                   vbOKOnly Or vbExclamation, _
                   "Range Error"
            Field.FormControl = ValString(TableVal)
            CheckRanges = False
        End If

    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' CheckRanges

Private Function ValidateWholeNumber( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateWholeNumber"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Dim Whole1 As Long: Whole1 = CInt(Evaluate(Field.ValidationFormula1))
    If Err.Number <> 0 Then Whole1 = 0
    On Error GoTo ErrorHandler

    On Error Resume Next
    Dim Whole2 As Long: Whole2 = CInt(Evaluate(Field.ValidationFormula2))
    If Err.Number <> 0 Then Whole2 = 0
    On Error GoTo ErrorHandler
    
    Dim TableVal As Long: TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
    
    On Error Resume Next
    Dim FormVal As Long: FormVal = Field.FormControl
    If Err.Number <> 0 Then
        MsgBox "Cell " & Field.Name & " must be a whole number", _
               vbOKOnly Or vbCritical, "Whole Number"
        On Error GoTo ErrorHandler
    End If
    
    ValidateWholeNumber = CheckRanges(Whole1, Whole2, FormVal, TableVal, Field)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateWholeNumber

Private Function ValidateDecimal( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateDecimal"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Dim Dec1 As Long: Dec1 = CDbl(Evaluate(Field.ValidationFormula1))
    If Err.Number <> 0 Then Dec1 = 0
    On Error GoTo ErrorHandler

    On Error Resume Next
    Dim Dec2 As Long: Dec2 = CDbl(Evaluate(Field.ValidationFormula2))
    If Err.Number <> 0 Then Dec2 = 0
    On Error GoTo ErrorHandler
    
    Dim TableVal As Double: TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
    
    On Error Resume Next
    Dim FormVal As Double: FormVal = Field.FormControl
    If Err.Number <> 0 Then
        MsgBox "Cell " & Field.Name & " must be a number", _
               vbOKOnly Or vbCritical, "Decimal Number"
        On Error GoTo ErrorHandler
    End If
    
    ValidateDecimal = CheckRanges(Dec1, Dec2, FormVal, TableVal, Field)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateDecimal

Private Function ValidateList( _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateList"
    On Error GoTo ErrorHandler
    
    Dim FormVal As Variant: FormVal = Field.FormControl
    
    If InScope(Field.ValidationList, FormVal, RoutineName) Then
        ValidateList = True
    Else
        MsgBox "The value in " & _
               Field.HeaderText & _
               " is not found in the validation list." & vbCrLf _
             & "Correct the value and try again.", _
               vbOKOnly Or vbCritical, _
               "Validation List Error"
    
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateList

Private Function ValidateDate( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateDate"
    On Error GoTo ErrorHandler
    
    Dim FormVal As Date
    On Error Resume Next
    FormVal = Field.FormControl
    If Err.Number <> 0 Or IsError(FormVal) Then
        FormVal = Empty
    End If
    On Error GoTo ErrorHandler
    If FormVal = 0 And Field.IgnoreBlank Then
        ValidateDate = True
        Exit Function
    End If
    
    On Error Resume Next
    Dim Date1 As Variant: Date1 = CDate(Field.ValidationFormula1)
    If Err.Number <> 0 Then Date1 = Empty
    On Error GoTo ErrorHandler

    On Error Resume Next
    Dim Date2 As Variant: Date2 = CDate(Field.ValidationFormula2)
    If Err.Number <> 0 Then Date2 = Empty
    On Error GoTo ErrorHandler
    
    Dim TableVal As Date: TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
    If TableVal = 0 Then TableVal = Empty
    
    On Error Resume Next
    If Err.Number <> 0 Then
        If FormVal = 0 Then
            ' Do nothing if the form value is blank
            ' An empty date or a zero is a "date"
        Else
            MsgBox "Cell " & Field.Name & " must be a date", _
                   vbOKOnly Or vbCritical, "Date Error"
        End If
    End If
    On Error GoTo ErrorHandler
    
    ValidateDate = CheckRanges(Date1, Date2, FormVal, TableVal, Field)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateDate

Private Function ValidateTime( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateTime"
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Dim Time1 As Date: Time1 = CDate(Evaluate(Field.ValidationFormula1))
    If Err.Number <> 0 Then Time1 = 0
    On Error GoTo ErrorHandler

    On Error Resume Next
    Dim Time2 As Date: Time2 = CDate(Evaluate(Field.ValidationFormula2))
    If Err.Number <> 0 Then Time2 = 0
    On Error GoTo ErrorHandler
    
    Dim FormVal As Date
    Dim TableVal As Date: TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
    
    On Error Resume Next
    FormVal = Field.FormControl
    If Err.Number <> 0 Then MsgBox "Cell " & Field.Name & " must be a date", _
       vbOKOnly Or vbCritical, "Time Error"
    On Error GoTo ErrorHandler
    
    ValidateTime = CheckRanges(Time1, Time2, FormVal, TableVal, Field)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateTime

Private Function ValidateTextLength( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateTextLength"
    On Error GoTo ErrorHandler
    
    Dim Lgth As Long: Lgth = CStr(Field.ValidationFormula1)

    On Error Resume Next
    Dim FormVal As String
    FormVal = Field.FormControl
    If Err.Number <> 0 Then _
       MsgBox "Cell " & Field.HeaderText & " must be a string ", _
       vbOKOnly Or vbCritical, "String Length Error"
    On Error GoTo ErrorHandler
    
    If Len(FormVal) = 0 And Field.IgnoreBlank Then
        ValidateTextLength = True
        Exit Function
    End If
    
    If Len(FormVal) = Lgth Then
        ValidateTextLength = True
    Else
        MsgBox Field.Name & " must be a string of length " & _
               Lgth, _
               vbOKOnly Or vbExclamation, _
               "String Length Error"
        Dim TableVal As String
        TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
        Field.FormControl = TimeFormat(TableVal)
        ValidateTextLength = False
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateTextLength

Private Function ValidateCustom( _
        ByVal Tbl As TableManager.TableClass, _
        ByVal Field As Variant _
        ) As Boolean

    ' Return True if value is validated
    
    Const RoutineName As String = Module_Name & "ValidateCustom"
    On Error GoTo ErrorHandler
    
    '    Dim FormVal As Variant: FormVal = Field.FormControl
    Dim TableVal As Variant: TableVal = Tbl.DBRange(Tbl.DBRow, Tbl.SelectedDBCol(Field.HeaderText))
    
    On Error Resume Next
    Dim ValForm1 As Variant: ValForm1 = Evaluate(Field.ValidationFormula1)
    If Err.Number <> 0 Then ValForm1 = vbNullString
    On Error GoTo ErrorHandler

    If ValForm1 Then
        ValidateCustom = True
    Else
        ValidateCustom = False
        Field.FormControl = ValString(TableVal)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' ValidateList

Public Sub PopulateForm( _
       ByVal Tbl As TableManager.TableClass, _
       ByVal Modulename As String)

    Const RoutineName As String = Module_Name & "PopulateForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)

    On Error GoTo ErrorHandler

    Dim Field As TableManager.CellClass
    Dim DBRange As Range: Set DBRange = Tbl.DBRange
    Dim DBRow As Long: DBRow = Tbl.DBRow
    Dim DBCol As Long
    Dim I As Long

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)
        DBCol = Tbl.SelectedDBCol(Field.HeaderText)
        If DBCol = 0 Then
            Err.Raise 1, "FormClass.PopulateForm", "Fatal error. HeaderText not found."
        End If

        Field.ControlValue = DBRange(DBRow, DBCol)
        
        Select Case Left$(Field.FormControl.Name, 3)
        Case "lbl":                              ' Do nothing
        Case "val": Field.FormControl.Caption = DBRange(DBRow, DBCol)
        Case "fld": Field.FormControl.Text = DBRange(DBRow, DBCol)
        Case "cmb": Field.FormControl.Text = DBRange(DBRow, DBCol)
        Case "whl": Field.FormControl.Text = DBRange(DBRow, DBCol)
        Case "dat": Field.FormControl.Text = DBRange(DBRow, DBCol)
        Case Else
            MsgBox "This is an illegal field type: " & _
                   Left$(Field.FormControl.Name, 3), _
                   vbOKOnly Or vbExclamation, "Illegal Field Type"

        End Select
        
    Next I
    
    TableManager.TurnOffCellDescriptions Tbl, Modulename
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' PopulateForm

Public Sub ClearForm( _
       ByVal Tbl As TableManager.TableClass, _
       ByVal Modulename As String)

    Const RoutineName As String = Module_Name & "ClearForm"
    Debug.Assert InScope(ModuleList, Modulename, RoutineName)

    On Error GoTo ErrorHandler

    Dim Field As TableManager.CellClass
    Dim I As Long

    For I = 0 To Tbl.CellCount - 1
        Set Field = Tbl.TableCells.Item(I, Module_Name)

        Select Case Left$(Field.FormControl.Name, 3)
        Case "lbl":                              ' Do nothing
        Case "val": Field.FormControl.Caption = vbNullString
        Case "fld": Field.FormControl.Text = vbNullString
        Case "cmb": Field.FormControl.Text = vbNullString
        Case "whl": Field.FormControl.Text = vbNullString
        Case "dat": Field.FormControl.Text = vbNullString
        Case Else
            MsgBox _
        "This is an illegal field type: " & Left$(Field.FormControl.Name, 3), _
                                            vbOKOnly Or vbExclamation, "Illegal Field Type"

        End Select
        
    Next I
    
    TableManager.TurnOffCellDescriptions Tbl, Modulename
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' ClearForm


