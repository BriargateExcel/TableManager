Attribute VB_Name = "XLAM_Module"
Option Explicit

Private Const Module_Name As String = "XLAM_Module."

Private Init As Boolean
Private pMainWorkbook As Workbook
Private pMainVBAProject As VBProject

Private pFormNameToShow As String

Public Sub FormNameToShow(ByVal FormNameToShow As String)
    pFormNameToShow = FormNameToShow
End Sub

Public Function GetFormName() As String
    GetFormName = pFormNameToShow
End Function

Public Function MainVBProject() As VBProject
    Set MainVBProject = pMainWorkbook.VBProject '(ThisWorkbook.Name)
End Function

Public Sub AutoOpen(ByVal WkBk As Workbook)
    
    Dim Sht As Worksheet
    Dim Tbl As ListObject
    Dim WkSht As TableManager.WorksheetClass
    
    Const RoutineName As String = Module_Name & "AutoOpen"
    On Error GoTo ErrorHandler
    
    Init = True
    If Not CheckForVBAProjectAccessEnabled(WkBk.Name) Then
        MsgBox "You must set the project access for the " & _
            "TableManager Add-In to work", _
            vbOKOnly Or vbCritical, _
            "Project Access"
    End If
    
    Set pMainWorkbook = WkBk
    Set pMainVBAProject = WkBk.VBProject

    
    
    Dim VBComp As Object
    For Each VBComp In WkBk.VBProject.VBComponents
        If VBComp.Type = vbext_ct_MSForm And _
            Left$(VBComp.Name, 8) = "UserForm" _
        Then
            WkBk.VBProject.VBComponents.Remove VBComp
        End If
    Next VBComp
    
    TableManager.WorksheetSetNewClass Module_Name
    TableManager.TableSetNewClass Module_Name
    
    For Each Sht In WkBk.Worksheets
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

End Sub ' AutoOpen

Public Sub BuildTable( _
    ByVal WS As TableManager.WorksheetClass, _
    ByVal TblObj As ListObject)
    
    Dim Tbl As Variant
    Dim Frm As TableManager.FormClass
    
    Const RoutineName As String = Module_Name & "BuildTable"
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
    
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub ' BuildTable

Public Function MainWorkbook() As Workbook
    Set MainWorkbook = pMainWorkbook
End Function

Public Function Initializing() As Boolean
    Initializing = Init
End Function ' Initializing

