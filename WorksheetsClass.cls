VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableManager.Worksheets")

Option Explicit

Private Const Module_Name As String = "WorksheetsClass."

Private pAllSheets As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllSheets = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.", "WorksheetRoutines.")
End Function                                     ' ModuleList

Public Property Get Count() As Long: Count = TableCount(pAllSheets.Count): End Property

'@DefaultMember
Public Property Get Item( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String) As WorksheetClass
    'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName As String = Module_Name & "Get_Item"
    Debug.Assert InScope(ModuleList, ModuleName)
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set Item = pAllSheets.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllSheets(vIndex)
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
       ByVal Sht As WorksheetClass, _
       ByVal ModuleName As String)

    Dim Evt As EventClass
    
    Const RoutineName As String = Module_Name & "Add"
    Debug.Assert Initializing
    Debug.Assert InScope(ModuleList, ModuleName)

    On Error GoTo ErrorHandler
    
    Set Evt = New EventClass
    
    If Sht.Name <> vbNullString Then
        pAllSheets.Add Sht.Name, Sht
        Set Evt.SheetEvent = Sht.Worksheet
        Set pAllSheets.Item(Sht.Worksheet.Name).WorksheetEvent.SheetEvent = Sht.Worksheet
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
    Debug.Assert InScope(ModuleList, ModuleName)
    
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        WorksheetSetNothing Module_Name
        WorksheetSetNewDict Module_Name
    Else
        If Not WorksheetExists(vIndex, Module_Name) Then Err.Raise 9
        pAllSheets.Remove vIndex
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Remove

Public Function Exists( _
       ByVal vIndex As Variant, _
       ByVal ModuleName As String _
       ) As Boolean
    ' Used in TableRoutines
    
    Const RoutineName As String = Module_Name & "Exists"
    On Error GoTo ErrorHandler
    
    Debug.Assert InScope(ModuleList, ModuleName)
    
    Exists = pAllSheets.Exists(vIndex)

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' Exists


