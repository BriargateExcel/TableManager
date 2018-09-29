VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Tables")

Option Explicit

Private Const Module_Name As String = "TablesClass."

Private pAllTables As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllTables = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("TableRoutines.")
End Function                                     ' ModuleList

Public Property Get Count() As Long: Count = pAllTables.Count: End Property

'@DefaultMember
Public Property Get Item( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String _
       ) As TableClass
    'Attribute Item.VB_UserMemId = 0

    Const RoutineName As String = Module_Name & "Get Item"
    On Error Resume Next
    
    Debug.Assert InScope(ModuleList, ModuleName)
        
    Set Item = pAllTables.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllTables(vIndex)
    If Err.Number <> 0 Then
        MsgBox "The Form for this table does not exist: " & CStr(vIndex) & ". Rebuild the Forms.", _
               vbOKOnly Or vbExclamation, _
               "Form Does Not Exist"
        Set Item = Nothing
        Exit Property
    End If
    On Error GoTo 0
    
    Debug.Assert Not Item Is Nothing

    '@Ignore LineLabelNotUsed
Done:
    Exit Property
    '@Ignore LineLabelNotUsed
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property

Public Sub Add( _
       ByVal Tbl As TableClass, _
       ByVal ModuleName As String)
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Add"
    On Error GoTo ErrorHandler
           
    Debug.Assert InScope(ModuleList, ModuleName)
    
    If Tbl.Name <> vbNullString Then pAllTables.Add Tbl.Name, Tbl
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Add

Public Function Exists( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String _
       ) As Boolean
    
    Const RoutineName As String = Module_Name & "Exists"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Exists = pAllTables.Exists(vIndex)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' Exists

Public Sub Remove( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    Debug.Assert InScope(ModuleList, ModuleName)
    
    If CStr(vIndex) = "*" Then
        Set pAllTables = Nothing
        Set pAllTables = New Scripting.Dictionary
    Else
        If Not pAllTables.Exists(vIndex) Then Err.Raise 9
        pAllTables.Remove vIndex
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Remove


