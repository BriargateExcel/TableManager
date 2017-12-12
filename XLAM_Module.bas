Attribute VB_Name = "XLAM_Module"
Option Explicit

' TODO Develop my own VBA error handling technique

Private Const Module_Name As String = "XLAM_Module."

Private Init As Boolean
Private pMainWorkbook As Workbook

Private Sub ExtendDataValidationDownTable(ByVal Tbl As TableManager.TableClass)
    
    Dim I As Long
    Dim CopyRange As Range
    
    For I = 1 To Tbl.CellCount
        Set CopyRange = Tbl.ColumnRange(I)
        CopyRange(1, 1).Copy
        CopyRange.PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next I
    
    Application.CutCopyMode = False

End Sub

Public Function MainWorkbook() As Workbook
    Set MainWorkbook = pMainWorkbook
End Function

Public Sub AutoOpen(ByVal Wkbk As Workbook)
    
    Dim Sht As Worksheet
    Dim Tbl As ListObject
    Dim UserFrm As Object
    Dim WkSht As TableManager.WorksheetClass
    
    Const RoutineName As String = Module_Name & "AutoOpen"
    On Error GoTo ErrorHandler
    
    Init = True
    Set pMainWorkbook = Wkbk
    
    If Not CheckForVBAProjectAccessEnabled(ThisWorkbook.Name) Then
        MsgBox "You must set the project access for the " & _
               "TableManager Add-In to work", _
               vbOKOnly Or vbCritical, _
               "Project Access"
    
    End If
    
    For Each UserFrm In Application.ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm And _
           Left$(UserFrm.Name, 8) = "UserForm" _
           Then
            Application.ThisWorkbook.VBProject.VBComponents.Remove UserFrm
        End If
    Next UserFrm
    
    TableManager.WorksheetSetNewClass Module_Name
    TableManager.TableSetNewClass Module_Name
    
    For Each Sht In MainWorkbook.Worksheets
        Set WkSht = New TableManager.WorksheetClass
        Set WkSht.Worksheet = Sht
        WkSht.Name = Sht.Name
        
        For Each Tbl In Sht.ListObjects
            BuildTable WkSht, Tbl
        Next Tbl
        
        TableManager.WorksheetAdd WkSht, Module_Name
    Next Sht
    
    DoEvents
    
    Init = False

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub                                          ' AutoOpen

Public Function Initializing() As Boolean
    Initializing = Init
End Function                                     ' Initializing


