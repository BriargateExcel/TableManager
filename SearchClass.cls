VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("TableAccess.SearchClass")
' Class used to gather (Set) and return (Get) Table data for all tables

Private Ary As Variant

Public Sub SetArray( _
       ByVal SearchTable As TableClass, _
       ByVal RowNum As Long)
    
    ' Assumes all arrays start at 1
    
    Dim I As Long
    Dim J As Long
    ReDim Ary(UBound(SearchTable.Headers, 2))
    For I = 1 To UBound(SearchTable.Headers, 2)
        Ary(I) = SearchTable.Body(RowNum, I)
        J = J + 1
    Next I
End Sub

Public Function GetArray() As Variant
    GetArray = Ary
End Function


