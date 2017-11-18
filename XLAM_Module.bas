Attribute VB_Name = "XLAM_Module"
Option Explicit

Private Const Module_Name = "XLAM_Module."

Private Init As Boolean
Private pMainWorkbook As Workbook

Public Function MainWorkbook() As Workbook
    Set MainWorkbook = pMainWorkbook
End Function

Public Sub BuildTable( _
    ByVal WS As TableManager.WorksheetClass, _
    ByVal TblObj As ListObject)
    
    Dim Tbl As Variant
    Dim Frm As TableManager.FormClass
    
    Const RoutineName = Module_Name & "BuildTable"
    On Error GoTo ErrorHandler
    
    ' Gather the table data
    Set Tbl = New TableManager.TableClass
    Tbl.Name = TblObj.Name
    Set Tbl.Table = TblObj
    If Tbl.CollectTableData(WS, Tbl) Then
        Set Frm = New TableManager.FormClass
        TableManager.TableAdd Tbl, Module_Name
        
        Set Frm.FormObj = Frm.BuildForm(Tbl)
        Set Tbl.Form = Frm
    End If
    
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildTable

Public Sub AutoOpen(ByVal WkBk As Workbook)
    
    Dim Sht As Worksheet
    Dim Tbl As ListObject
    Dim UserFrm As Object
    Dim WkSht As TableManager.WorksheetClass
    
    Const RoutineName = Module_Name & "AutoOpen"
    On Error GoTo ErrorHandler
    
    Init = True
    Set pMainWorkbook = WkBk
    
    CheckForVBAProjectAccessEnabled (ThisWorkbook.Name)
    
    For Each UserFrm In Application.ThisWorkbook.VBProject.VBComponents
        If UserFrm.Type = vbext_ct_MSForm And _
            Left(UserFrm.Name, 8) = "UserForm" _
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

Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName

End Sub ' AutoOpen

Public Function Initializing() As Boolean
    Initializing = Init
End Function ' Initializing
