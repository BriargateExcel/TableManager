VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlainDataBaseForm 
   Caption         =   "Save and Restore Table Data"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "PlainDataBaseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlainDataBaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TableManager.DataBase")

Option Explicit

Private Const Module_Name As String = "PlainDataBaseForm."

Private pTableObj As TableClass

Public Sub SetTable(ByVal Tbl As TableClass)
    Set pTableObj = Tbl
End Sub

Private Sub CopyButton_Click()
    OutputTable GetMainWorkbook, Module_Name
    Me.Hide
End Sub

Private Sub FetchButton_Click()
    InputTable GetMainWorkbook, Module_Name
    Me.Hide
End Sub

Private Sub ChangeFileButton_Click()
    ChangeFile Module_Name
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    Me.Hide

    FileNameBox.Text = vbNullString
End Sub

Private Sub UserForm_Activate()
    FileNameBox.Text = GetFullFileName(GetMainWorkbook, ActiveCellTableName)
    CenterMe Me
End Sub

