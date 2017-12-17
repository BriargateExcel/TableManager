VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Module_Name As String = "TablesClass."

Private pAllTables As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllTables = New Scripting.Dictionary
End Sub                                          ' Class_Initialize

Public Property Get Count() As Long: Count = pAllTables.Count: End Property

'@DefaultMember
Public Property Get Item(ByVal vIndex As Variant) As TableManager.TableClass
    'Attribute Item.VB_UserMemId = 0

    Const RoutineName As String = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllTables.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    On Error Resume Next
    Set Item = pAllTables(vIndex)
    If Err.Number <> 0 Then
        MsgBox "The Form for this table does not exist. Rebuild the Forms.", _
               vbOKOnly Or vbExclamation, _
               "Form Does Not Exist"
    End If
    On Error GoTo 0

    '@Ignore LineLabelNotUsed
Done:
    Exit Property
    '@Ignore LineLabelNotUsed
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property
Public Sub Add(ByVal Tbl As TableManager.TableClass)
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Add"
    On Error GoTo ErrorHandler
           
    If Tbl.Name <> vbNullString Then pAllTables.Add Tbl.Name, Tbl
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub                                          ' Add

Public Function Exists(ByVal vIndex As Variant) As Boolean
    Exists = pAllTables.Exists(vIndex)
End Function                                     ' Exists

Public Sub Remove(ByVal vIndex As Variant)
    
    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
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

