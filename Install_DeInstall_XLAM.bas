Attribute VB_Name = "Install_DeInstall_XLAM"
Option Explicit

Private Const Module_Name As String = "InitializeXLAM."

Private Const TableManager As String = "TableManager"

Private Sub InstallXLAM()
    
    Dim TableManagerFileName As String
    TableManagerFileName = TableManager & ".xlam"
    
    Dim TableManagerFullPath As String
    TableManagerFullPath = ThisWorkbook.Path & "\" & TableManagerFileName
    
    Dim TableManagerAddIn As Workbook
    Dim LastError As Long
    
    Set TableManagerAddIn = Workbooks.Open(TableManagerFullPath)
    AddIns(TableManager).Installed = True
        
    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisWorkbook.VBProject
    vbProj.References.AddFromFile (TableManagerFullPath)
    
End Sub                                          ' InstallXLAM

Private Sub DeInstallXLAM()
    ' Use this to eliminate the reference to TableManager from this VBAProject

    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisWorkbook.VBProject
    
    Dim Ref As Reference
    For Each Ref In vbProj.References
        If Ref.Name = TableManager Then
            vbProj.References.Remove Ref
        End If
    Next Ref
    
    Dim TableManagerFileName As String
    TableManagerFileName = TableManager & ".xlam"
    Workbooks(TableManagerFileName).Close
    
End Sub                                          ' DeInstallXLAM

