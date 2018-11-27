VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Tables")

Option Explicit

Private Const Module_Name As String = "TableClass."

Private pTableName As String
Private pWkSht As Worksheet
Private pWorkbook As Workbook
Private pUserForms As Collection
Private pForm As FormClass
Private pCells As CellsClass
Private pTable As ListObject

Private pMaxHeader As Single
Private pMaxField As Single
Private pWrapCount As Single
Private pLocked As Boolean
Private pTarget As Range

Private Type TableType
    ' Headers
    '   The table's HeaderRowRange
    '   2D array
    '   Only has row 1 (first parameter)
    '   The column names are in the column (second) parameter of the array
    '   If it exists, Row 0 is not used
    '   If it exists, Column 0 is not used
    ' Body
    '   The table's DataBodyRange
    '   Each table row is designated by the first (row) parameter of the array
    '   Each column is designated by the second (column) parameter of the array
    '   If it exists, Row 0 is not used
    '   If it exists, Column 0 is not used
    ' Valid
    '   "Valid" if TableType valid
    '   Error message if TableType invalid
    Headers As Variant
    Body As Variant
    Valid As String
End Type

Private pDataTable As TableType

Private Type ColumnDesignatorType
    ColumnName As String
    ColumnNumber As Long
End Type

Private Sub Class_Initialize()
    Set pCells = New CellsClass
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("TableRoutines.")
End Function                                     ' ModuleList

Public Property Get ActiveRow() As Long
    ' Used in EventClass
    ActiveRow = Me.DBRow
End Property

Public Property Get ActiveTarget() As Range
    ' Used in EventClass, TableClass
    Set ActiveTarget = pTarget
End Property

Public Property Set ActiveTarget(ByVal Target As Range)
    ' Used in EventClass

    Target.Select
    Set pTarget = Target
End Property

Public Property Get DBRange() As Range
    ' Used in FormRoutines, TableClass, TableRoutines, CalendarBuilder
    Set DBRange = pWkSht.ListObjects(pTableName).DataBodyRange
End Property

Public Property Get UserForms() As Workbook
    Set UserForms = pUserForms
End Property

Public Property Set UserForms(ByVal vbl As Workbook)
    Set pUserForms = vbl
End Property

Public Property Get DBColRange( _
       ByVal Tbl As TableClass, _
       ByVal ColumnName As String _
       ) As Range
    ' Used in TableRoutines, TableClass
    
    Dim ColNum As Long
    ColNum = Tbl.DBColNumber(ColumnName)
    If ColNum = 0 Then
        Err.Raise 1, "TableClass.DBColRange", "Fatal error. ColumnName not found."
    End If
    
    Set DBColRange = Tbl.Table.ListColumns(ColNum).DataBodyRange
End Property

Public Property Get DBColNumber(ByVal ColumnName As String) As Long
    ' Used in TableRoutines, TableClass
    
    On Error Resume Next
    DBColNumber = Application.WorksheetFunction.Match(ColumnName, Me.Table.HeaderRowRange, 0)
    If Err.Number <> 0 Then DBColNumber = 0
    On Error GoTo 0
End Property

Public Property Get DBRowNumber( _
       ByVal ColumnName As String, _
       ByVal Key As String _
       ) As Long
       
    ' Used in TableRoutines, TableClass
    
    Dim KeyRange As Range
    Set KeyRange = Me.DBColRange(Me, ColumnName)
    
    On Error Resume Next
    DBRowNumber = Application.WorksheetFunction.Match(Key, KeyRange, 0)
    If Err.Number <> 0 Then DBRowNumber = 0
    On Error GoTo 0
End Property

Public Property Get Form() As FormClass: Set Form = pForm: End Property

Public Property Set Form(ByVal Frm As FormClass): Set pForm = Frm: End Property

Public Property Get Locked() As Boolean: Locked = pLocked: End Property

Public Property Get MaxFieldLength() As Single: MaxFieldLength = pMaxField: End Property

Public Property Get MaxHeaderLength() As Single: MaxHeaderLength = pMaxHeader: End Property

Public Property Get Name() As String: Name = pTableName: End Property

Public Property Let Name(ByVal vbl As String): pTableName = vbl: End Property

Public Property Get NumColumns() As Long: NumColumns = pCells.Count: End Property

Public Property Get NumRows() As Long: NumRows = DBRange.Rows.Count: End Property

Public Property Get TableCells() As CellsClass: Set TableCells = pCells: End Property

Public Property Get Target() As Range: Set Target = pTarget: End Property

Public Property Set Target(ByVal vbl As Range): Set pTarget = vbl: End Property

Public Property Get Workbook() As Workbook: Set Workbook = pWorkbook: End Property

Public Property Get Worksheet() As Worksheet: Set Worksheet = pWkSht: End Property

Public Property Get WrapCount() As Long: WrapCount = pWrapCount: End Property

Public Property Get Table() As ListObject: Set Table = pTable: End Property

Public Property Set Table(ByVal Tbl As ListObject): Set pTable = Tbl: End Property

'@Ignore FunctionReturnValueNotUsed
Public Function CollectTableData( _
       ByVal Wkbk As Workbook, _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String _
       ) As Boolean
    ' Used in TableClass
       
    ' Call CollectCellData to collect all the type information for each column of the table
    
    Const RoutineName As String = Module_Name & "CollectTableData"
    On Error GoTo ErrorHandler
    
    Debug.Assert Initializing

    Debug.Assert InScope(ModuleList, ModuleName)
    
    pTableName = Tbl.Table.Name
    Set pWkSht = Wkbk.Worksheets(Tbl.Table.Parent.Name)
    Set pWorkbook = Wkbk
    pLocked = True
    Dim Element As Range
    Dim OneCell As CellClass
    
    For Each Element In pTable.HeaderRowRange
        Set OneCell = New CellClass
        OneCell.CollectCellData Tbl, Element, Module_Name
        pCells.Add OneCell, Module_Name
        If OneCell.ColumnWidth > pMaxField Then pMaxField = OneCell.ColumnWidth
        If OneCell.Wrap Then pWrapCount = pWrapCount + 1
        If OneCell.Length > pMaxHeader Then pMaxHeader = OneCell.Length
        If Not OneCell.Locked Then pLocked = False
    Next Element
    
    If pLocked Then
        MsgBox "All cells in the " & _
               pTableName & _
               " table are locked. No form created.", _
               vbOKOnly Or vbCritical, _
               "All Cells Locked"

    End If
    
    LoadTable Wkbk, Tbl.Table, pDataTable

    CollectTableData = Not pLocked

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' CollectTableData

Public Function FirstCell() As Range
    Const RoutineName As String = Module_Name & "FirstCell"
    On Error GoTo ErrorHandler
    ' Used in EventClass, TableClass, TableRoutines
    
    On Error Resume Next
    Set FirstCell = pTable.DataBodyRange(1, 1)
    On Error GoTo 0
    If Err.Number = 91 Then
        Set FirstCell = pTable.HeaderRowRange(2, 1)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' FirstCell

Public Function LastCell() As Range
    ' Used in TableClass, EventClass
    Const RoutineName As String = Module_Name & "LastCell"
    On Error GoTo ErrorHandler
    
    Set LastCell = pTable.DataBodyRange(pTable.ListRows.Count, 1)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' LastCell

Public Function LastRow() As Long
    LastRow = pTarget.ListObject.ListRows.Count
End Function                                     ' LastRow

Public Sub NextRow()

    ' Called in EventClass
    
    Dim DBRow As Long
    Dim I As Long
    
    DBRow = Me.DBRow
    
    If DBRow < pTarget.ListObject.ListRows.Count Then
        TurnOnCellDescriptions Me, Module_Name
        For I = DBRow To DBRange.Rows.Count
            Set pTarget = pTarget.Offset(1)
            If Not pTarget.Rows.Hidden Then Exit For
        Next I
        TurnOffCellDescriptions Me, Module_Name
    Else
        MsgBox "Cannot advance to the next row; already at last row", _
               vbOKOnly Or vbExclamation, "Already at End of Table"
    End If
End Sub                                          ' NextRow

Public Sub PreviousRow()

    ' Called in EventClass
    
    Dim DBRow As Long
    Dim I As Long
    
    DBRow = Me.DBRow
    
    If DBRow > 1 Then
        TurnOnCellDescriptions Me, Module_Name
        For I = DBRow To DBRange.Rows.Count
            Set pTarget = pTarget.Offset(-1)
            If Not pTarget.Rows.Hidden Then Exit For
        Next I
        TurnOffCellDescriptions Me, Module_Name
    Else
        MsgBox "Cannot move to the previous row; already at first row", _
               vbOKOnly Or vbExclamation, _
               "Already at Beginning of Table"
    End If
End Sub                                          ' PreviousRow

Public Function DBRow() As Long

    ' Called in TableClass, FormRoutines, TableRoutines,
    Dim THRow As Long: THRow = Me.ActiveTarget.ListObject.HeaderRowRange.Row
    Dim TRow As Long: TRow = Me.ActiveTarget.Row
    DBRow = TRow - THRow
End Function                                     ' DBRow

Public Function SelectedDBCol(ByVal ColumnName As String) As Long

    ' Called in FormRoutines, TableRoutines
    
    Dim Rng As Range
    Dim vbl As Variant
    
    Set Rng = Me.ActiveTarget.ListObject.HeaderRowRange
    On Error Resume Next
    vbl = Application.WorksheetFunction.Match(ColumnName, Rng, 0)
    If Err.Number <> 0 Then
        SelectedDBCol = 0
        Exit Function
    End If
    On Error GoTo 0
    SelectedDBCol = vbl
End Function                                     ' SelectedDBCol

Public Function ColumnRange(ByVal ColNum As Long) As Range

    ' Called in TableRoutines
    
    Set ColumnRange = pWkSht.Range(DBRange(1, ColNum), DBRange(NumRows, ColNum))
End Function

Public Function CellCount() As Long: CellCount = pCells.Count: End Function
' Called in FormClass, TableClass, FormRoutines, TableRoutines

Public Property Get Headers() As Variant
    Headers = pDataTable.Headers
End Property

Public Property Get Body() As Variant
    Body = pDataTable.Body
End Property

Public Property Let Body(ByVal vbl As Variant)
    pDataTable.Body = vbl
End Property

Public Function GetData( _
       Optional ByVal RowDesignator As Variant = "Empty", _
       Optional ByVal ColumnDesignator As Variant = "Empty", _
       Optional ByVal ColumnFilter As String = "Empty" _
       ) As Variant
       
    ' Purpose
    '   Return the subset of SearchTable as specified in the other parameters
    ' Assumptions
    
    ' Future:
    ' Add provisions for multiple filters; "And" only; no "Or"
    ' ColumnFilter becomes a parameter array
    ' Alternately, use this routine for the first filter then pass
    ' this routine's output to another routine for the next filter
    
    ' If RowDesignator is a number, that's the row to return
    ' If RowDesignator is "Empty" return the rows specified by ColumnFilter
    
    ' If ColumnDesignator has a value, that's the column to return
    ' If ColumnDesignator = "Empty", select the entire row
    '
    ' If ColumnFilter contains "=", there's only one row to return
    ' If ColumnFilter contains "<>", "<', ">", "<=", or ">="
    '   there are (potentially) multiple rows to return
    '
    ' Symbology for the table:
    ' RowDesignator (RD) can be numeric (N) or "Empty" (E)
    ' ColumnDesignator (CD) can be specified (S) or "Empty" (E)
    ' ColumnFilter (CF) can result in
    '   0 hits (0)
    '   1 hit (1)
    '   Multiple hits (M)
    '   "Empty" (E)
    '
    ' Table below:
    ' RD CD CF Result
    ' NS0 Makes no sense
    ' NS1 Makes no sense
    ' NSM Makes no sense
    ' NSE Single value
    ' NE0 Makes no sense
    ' NE1 Makes no sense
    ' NEM Makes no sense
    ' NEE One row
    ' ES0 No data
    ' ES1 Single value
    ' ESM One column
    ' ESE One column
    ' EE0 No data
    ' EE1 Single value
    ' EEM Multiple rows, All columns
    ' EEE All Rows, All columns - the entire table
    '
    ' The output can be
    '   No data
    '       ColumnFilter evaluates to 0 rows (regardless of RowDesignator and ColumDesignator value)
    '   A single value
    '       RowDesignator is numeric and ColumnDesignator is specified and ColumnFilter is empty
    '   A single row
    '       RowDesignator is numeric and ColumnDesignator is "Empty" and ColumnFilter is empty
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to a single row
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to a single row
    '   A single column
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter is "Empty"
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to multiple rows
    '   An array of rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to multiple rows
    '   All rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter is "Empty"
    '   Makes no sense to specify a row and a filter
    '       RowDesignator is numeric and ColumnFilter <> "Empty"
    '
    ' Error messages:
    '   "Error Table" if the SearchTable is invalid
    '   "Error Row Designator" if Rowdesignator is invalid
    '   "Error RowDesignator Out of Range"
    '   "Error Column Designator" if ColumnDesignator is invalid
    '   "Error Column Name Not Found"
    '   "Error Filter" if ColumnFilter is invalid
    '   "Error No Data" if ColumnFilter eliminates all the rows
    '   "Error Can't have a specific row and a column filter"
    '   Note that the calling routine need only check for "Error"
    '       to determine if there's an error and need only
    '       go deeper if necessary
    
    ' Start of code
    '
    Const Routine_Name As String = Module_Name & "GetData"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(pDataTable.Valid, 5) = "Error" Then
        pDataTable.Valid = "Error Table"
        Exit Function
    End If
    
    ' Verify that RowDesignator is valid
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(pDataTable, RowDesignator)
    
    Select Case Left$(TempRowDesignator, 5)
    Case "Error"                                 ' Invalid RowDesignator
        GetData = "Error Row Designator"
        Exit Function
    Case "Empty"                                 ' Empty RowDesignator
        RowNumber = 0
    Case Else                                    ' Conclude that RowDesignator must be a number
        If ColumnFilter <> "Empty" Then
            GetData = "Error Can't have a specific row and a column filter"
            Exit Function
        End If
        RowNumber = TempRowDesignator
    End Select
    
    ' Verify that ColumnDesignator is valid
    Dim ColumnNumber As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(pDataTable, ColumnDesignator)
    Select Case Left$(TempColumnDesignator.ColumnName, 5)
    Case "Error"                                 ' Invalid ColumnDesignator
        GetData = "Error Column Designator"
        Exit Function
    Case "Empty"                                 ' Empty ColumnDesignator
        ColumnNumber = 0
    Case Else
        ColumnNumber = TempColumnDesignator.ColumnNumber
    End Select
    
    ' Verify that ColumnFilter is valid and set up the SearchTable
    Dim ThisSearchTable As TableType
    ThisSearchTable = ValidFilter(pDataTable, ColumnFilter)
    
    Select Case Left$(ThisSearchTable.Valid, 5)
    Case "Error"
        GetData = ThisSearchTable.Valid
        Exit Function
    Case "Empty"                                 ' An empty filter means return all the rows
        ThisSearchTable = pDataTable
    Case Else
        ' ThisSearchTable is already set up in the ValidFilter call
    End Select
    
    Dim RowCount As Long
    RowCount = UBound(ThisSearchTable.Body, 1)
    
    ' SearchTable, RowDesignator, ColumnDesignator, and ColumnFilter are all valid
    If RowDesignator <> "Empty" Then             ' Valid RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                ' The case where RowDesignator <> "Empty" and ColumnFilter <> "Empty" cannot exist
                GetData = "Error Can't have a specific row and a column filter"
            Else                                 ' 1 row, 1 column, empty filter; one cell
                GetData = ThisSearchTable.Body(RowNumber, ColumnNumber)
                '                pDataTable.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            If ColumnFilter <> "Empty" Then
                ' The case where RowDesignator <> "Empty" and ColumnFilter <> "Empty" cannot exist
            Else                                 ' 1 row, unspecified column, empty filter; entire row
                GetData = GetRow(ThisSearchTable, RowDesignator)
                '                pDataTable.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    Else                                         ' Empty RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, 1 column, filter=0 rows; no data
                    GetData = "Empty"
                Case 1                           ' unspecified row, 1 column, filter=1 row; one cell
                    GetData = GetColumn(ThisSearchTable, ColumnNumber)
                    '                    pDataTable.Valid = "Valid"
                Case Else                        ' unspecified row, 1 column, filter=multiple rows; one column
                    GetData = GetColumn(ThisSearchTable, ColumnNumber)
                    '                    pDataTable.Valid = "Valid"
                End Select
            Else                                 ' empty row, one column, empty filter; one entire column
                GetData = GetColumn(ThisSearchTable, ColumnNumber)
                '                pDataTable.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, unspecified column, filter=0 rows; no data
                    GetData = "Empty"
                Case 1                           ' unspecified row, unspecified column, filter=1 row; one row, all columns
                    GetData = ThisSearchTable.Body
                Case Else                        ' unspecified row, unspecified column, filter=multiple rows; multiple rows, all columns
                    GetData = ThisSearchTable.Body
                End Select
            Else                                 ' empty row, empty column, empty filter; entire table
                GetData = ThisSearchTable.Body
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    End If                                       ' RowDesignator <> "Empty"
    
    '    pDataTable.Headers = ThisSearchTable.Headers

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Property Get Valid() As Boolean
    Valid = pDataTable.Valid
End Property

Private Property Let Valid(ByVal Val As Boolean)
    pDataTable.Valid = Val
End Property

Private Function GetRow( _
        SearchTable As TableType, _
        ByVal RowNum As Long _
        ) As Variant
        
    ' Purpose
    '   Returns the row designated by RowNum
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "GetRow"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        GetRow.Valid = "Error Table"
        Exit Function
    End If
    
    Dim Ary As Variant
    Dim NumberOfColumns As Long
    Dim LastError As Long
    On Error Resume Next                         ' UBound(SearchTable.Headers, 2) raises an error if there's only one column
    NumberOfColumns = UBound(SearchTable.Headers, 2)
    LastError = Err.Number
    On Error GoTo ErrorHandler
    
    If LastError = 0 Then
        ReDim Ary(1, NumberOfColumns)
        Dim I As Long
        For I = 1 To UBound(SearchTable.Headers, 2)
            Ary(1, I) = SearchTable.Body(RowNum, I)
        Next I
        
        GetRow = Ary
    Else
        GetRow = SearchTable.Body(RowNum, 1)
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function GetColumn( _
        SearchTable As TableType, _
        ByVal ColumnNum As Long _
        ) As Variant
        
    ' Purpose
    '   Returns the column designated by ColumnNum
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "GetColumn"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        GetColumn.Valid = "Error Table"
        Exit Function
    End If
    
    Dim Ary As Variant
    ReDim Ary(UBound(SearchTable.Body, 1), 1)
    
    Dim I As Long
    For I = 1 To UBound(SearchTable.Body, 1)
        Ary(I, 1) = SearchTable.Body(I, ColumnNum)
    Next I
    
    GetColumn = Ary
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidRowDesignator( _
        SearchTable As TableType, _
        ByVal RowDesignator As Variant _
        ) As Variant

    ' Purpose
    '   Determine if RowDesignator is valid
    '   Return the valid row number or "Empty" if valid
    '   Return ValidRowDesignator = "Error RowDesignator Out of Range" or
    '       "Error Row Designator" if invalid
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "ValidRowDesignator"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidRowDesignator.Valid = "Error Table"
        Exit Function
    End If
    
    If RowDesignator = "Empty" Then
        ' Empty is a valid entry; means return an entire column
        ValidRowDesignator = "Empty"
    Else
        Select Case VarType(RowDesignator)
        Case vbInteger, vbLong
            ' RowDesignator is numeric
            ' Verify that RowDesignator is in the range of the table's rows
            If RowDesignator >= 1 And RowDesignator <= UBound(SearchTable.Body, 1) Then
                ValidRowDesignator = RowDesignator
            Else
                ' RowDesignator is out of range
                ValidRowDesignator = "Error RowDesignator Out of Range"
            End If
        Case Else
            ' Erroneous RowDesignator data type
            ValidRowDesignator = "Error Row Designator"
        End Select
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidColumnDesignator( _
        SearchTable As TableType, _
        ByVal ColumnDesignator As Variant _
        ) As ColumnDesignatorType
    
    ' Purpose
    '   Determine if ColumnDesignator or "Empty" if valid
    '   Return the valid column name and number if valid
    '   Return ValidColumnDesignator.ColumnName = "Error Column Name Not Found" if invalid
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "ValidColumnDesignator"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidColumnDesignator.ColumnName = "Error Table"
        Exit Function
    End If
    
    Dim ColumnFound As Boolean
    Dim ColumnNumber As Long
    Select Case VarType(ColumnDesignator)
    Case vbInteger, vbLong
        If ColumnDesignator >= LBound(SearchTable.Headers, 2) And _
                                                              ColumnDesignator <= UBound(SearchTable.Headers, 2) Then
            ValidColumnDesignator.ColumnNumber = ColumnDesignator
        Else
            SearchTable.Valid = "Error Column Designator Out of Range"
        End If
    Case vbString
        If ColumnDesignator = "Empty" Then
            ' "Empty" is a valid entry; means return an entire row
            ' Set ColumnNumber to 0 as a flag
            ValidColumnDesignator.ColumnName = "Empty"
            ValidColumnDesignator.ColumnNumber = 0
        Else
            ColumnFound = False
            
            For ColumnNumber = 1 To UBound(SearchTable.Headers, 2)
                If SearchTable.Headers(1, ColumnNumber) = ColumnDesignator Then
                    ColumnFound = True
                    ValidColumnDesignator.ColumnName = ColumnDesignator
                    ValidColumnDesignator.ColumnNumber = ColumnNumber
                    Exit For
                End If
            Next ColumnNumber
            If Not ColumnFound Then
                ValidColumnDesignator.ColumnName = "Error Column Name Not Found"
                ValidColumnDesignator.ColumnNumber = 0
            End If
        End If
    Case Else
        SearchTable.Valid = "Error Invalid Column Name"
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidFilter( _
        SearchTable As TableType, _
        ByVal ColumnFilter As String _
        ) As TableType
        
    ' Purpose
    '   Determines if the ColumnFilter is valid
    '   If ColumnFilter valid, returns the filtered array
    '   If ColumnFilter is invalid, returns ValidFilter.Valid to an error message
    ' Assumptions
    ' The structure of ColumnFilter is column name
    '   then a boolean comparator then the operand.
    
    Const Routine_Name As String = Module_Name & "ValidFilter"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidFilter.Valid = "Error Table"
        Exit Function
    End If
    
    If ColumnFilter = "Empty" Then
        ' "Empty" is a valid ColumnFilter value
        ValidFilter = SearchTable
    Else
        ' Parse the ColumnFilter
        Dim I As Long
        Dim EndOfColumnName As Long
        Dim StartOfOperand As Long
        Dim PrevChar As String
        Dim ThisChar As String
        Dim NextChar As String
        Dim Operator As String
        Dim Operand As String
        Dim ColumnName As String
        Dim ColumnNumber As Long
        For I = 1 To Len(ColumnFilter) - 1
            ' Iterate through ColumnFilter looking for "=", "<", or ">"
            PrevChar = ThisChar
            ThisChar = Mid$(ColumnFilter, I, 1)
            NextChar = Mid$(ColumnFilter, I + 1, 1)
            
            Select Case ThisChar
            Case "="
                Operator = "="
                
                If PrevChar = " " Then
                    ColumnName = Mid$(ColumnFilter, 1, I - 2)
                Else
                    ColumnName = Mid$(ColumnFilter, 1, I - 1)
                End If
                
                Dim TempValidColumnDesignator As ColumnDesignatorType
                TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                If TempValidColumnDesignator.ColumnName = "Error" Then
                    ValidFilter.Valid = "Error filter"
                    Exit Function
                End If
                ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                If NextChar = " " Then
                    Operand = Mid$(ColumnFilter, I + 2, Len(ColumnFilter) - I - 1)
                Else
                    Operand = Mid$(ColumnFilter, I + 1, Len(ColumnFilter) - I)
                End If
                
            Case "<"                             ' ThisChar = "<"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                Select Case NextChar
                Case "="
                    Operator = "<="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case ">"
                    Operator = "<>"
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case Else                        ' NextChar <> "=" and NextChar <> ">"
                    Operator = "<"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End Select
            Case ">"                             ' ThisChar = ">"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                If NextChar = "=" Then
                    Operator = ">="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Else
                    Operator = ">"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
        
        Dim FilterCriteria As String
        FilterCriteria = " " & Operator & " " & """" & Operand & """"
        ValidFilter = SetUpSearchTable(FilterCriteria, ColumnNumber)
        If Left$(ValidFilter.Valid, 5) = "Error" Then
            Exit Function
        End If
    End If                                       ' ColumnFilter = "Empty"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function SetUpSearchTable( _
        ByVal FilterCriteria As String, _
        ByVal FilterColumnNumber As Long _
        ) As TableType
        
    ' Purpose
    '   Filters Searchtable according to the FilterCriteria
    '   If ColumnFilter valid, returns the filtered array
    '   If ColumnFilter is invalid, sets SetUpSearchTable.Valid to an error message
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "SetUpSearchTableColl"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(pDataTable.Valid, 5) = "Error" Then
        SetUpSearchTable.Valid = "Error Table"
        Exit Function
    End If
    
    Dim I As Long
    Dim SearchElement As SearchClass
    Dim Expression As String
    
    Dim SearchCollection As Collection
    Set SearchCollection = New Collection
    
    ' Put all the rows that match FilterCriteris into a collection
    For I = 1 To UBound(pDataTable.Body, 1)
        Expression = Chr$(34) & pDataTable.Body(I, FilterColumnNumber) & Chr$(34) & FilterCriteria
        If Evaluate(Expression) Then
            Set SearchElement = New SearchClass
            SearchElement.SetArray Me, I
            SearchCollection.Add SearchElement
        End If
    Next I
    
    If SearchCollection.Count = 0 Then
        SetUpSearchTable.Valid = "Error No data found"
        Exit Function
    End If
    
    Dim DataArray As TableType
    ReDim DataArray.Body(SearchCollection.Count, UBound(pDataTable.Headers, 2))
    ReDim DataArray.Headers(1, UBound(pDataTable.Headers, 2))
    DataArray.Valid = "Valid"
    
    DataArray.Headers = pDataTable.Headers
    
    Dim RowArray As Variant
    ReDim RowArray(UBound(pDataTable.Headers, 2))
    
    Dim J As Long
    
    ' Extract all the elements in the collection and put them into DataArray
    I = 1
    Dim ValidRow As Variant
    For Each ValidRow In SearchCollection
        RowArray = ValidRow.GetArray
        For J = 1 To UBound(pDataTable.Headers, 2)
            DataArray.Body(I, J) = RowArray(J)
        Next J
        I = I + 1
    Next ValidRow
    
    SetUpSearchTable = DataArray
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

'Public Sub GoToHome()
'    Set pTarget = Me.FirstCell
'End Sub

Private Sub LoadTable( _
        ByVal Wkbk As Workbook, _
        ByVal ExcelTable As ListObject, _
        ByRef DataTable As TableType)
       
    ' Called in TableClass
       
    ' Purpose
    ' Loads the contents of an Excel table (ExcelTable)
    '   into a TableType data structure (DataTable)
    ' Errors are reported via standard error handling logic
    ' Assumptions
    ' The structure of ColumnFilter is column name
    '   then a boolean comparator then the operand.

    Const Routine_Name As String = Module_Name & "LoadTable"
    On Error GoTo ErrorHandler
    
    Dim Sht As Worksheet
    Set Sht = Wkbk.Worksheets(ExcelTable.Parent.Name)
    
    Dim TableName As String
    TableName = ExcelTable.Name

    If Not IsArrayAllocated(DataTable.Body) Then
        DataTable.Headers = Sht.ListObjects(TableName).HeaderRowRange
        DataTable.Body = Sht.ListObjects(TableName).DataBodyRange
        DataTable.Valid = "Valid"
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Sub


