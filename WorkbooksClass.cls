VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbooksClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Workbooks")

Option Explicit

Private Const Module_Name As String = "WorkbooksClass."

Private pAllBooks As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllBooks = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.")
End Function                                     ' ModuleList

Public Property Get Count() As Long: Count = TableCount(pAllBooks.Count): End Property

'@DefaultMember
Public Property Get Item( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String) As WorkbookClass
    'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName As String = Module_Name & "Get_Item"
    Debug.Assert InScope(ModuleList, ModuleName)
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set Item = pAllBooks.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllBooks(vIndex)
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
       ByVal Wkbk As WorkbookClass, _
       ByVal ModuleName As String)

    Dim Evt As EventClass
    
    Const RoutineName As String = Module_Name & "Add"
    Debug.Assert Initializing
    Debug.Assert InScope(ModuleList, ModuleName)

    On Error GoTo ErrorHandler
    
    Set Evt = New EventClass
    
    If Wkbk.Name <> vbNullString Then
        pAllBooks.Add Wkbk.Name, Wkbk
        Set Evt.WorkbookEvent = Wkbk
        Set pAllBooks.Item(Wkbk.Name).WorkbookEvent.bookEvent = Wkbk
    End If
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub                                          ' Add

Public Sub Remove( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String)

    Const RoutineName As String = Module_Name & "Remove"
    Debug.Assert Initializing
    Debug.Assert InScope(ModuleList, ModuleName)
    
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        TableSetNothing Module_Name
        TableSetNewDict Module_Name
    Else
        If Not TableExists(vIndex, Module_Name) Then Err.Raise 9
        TableRemove vIndex, Module_Name
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Remove


