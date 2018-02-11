VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataBaseForm 
   Caption         =   "UserForm10"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "DataBaseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataBaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements I_CopyFetchFormView

Private Type pTView
    Model As CopyFetchClass
    IsCancelled As Boolean
    Source As String
    Destination As String
End Type

Private pThis As pTView

Private Sub DestinationBox_Change()
    DestinationBox.Text = pThis.Destination
End Sub

Public Property Let I_CopyFetchFormView_Source(ByVal Src As String)
    pThis.Source = Src
End Property

Public Property Let I_CopyFetchFormView_Destination(ByVal Dst As String)
    pThis.Destination = Dst
End Property

Private Sub CopyButton_Click()
    pThis.Model.CopyClicked = True
    Me.Hide
End Sub

Private Sub FetchButton_Click()
    pThis.Model.FetchClicked = True
    Me.Hide
End Sub

Private Function I_CopyFetchFormView_ShowDialog(ByVal viewModel As Object) As Boolean
    DestinationBox.Text = pThis.Destination
    SourceBox.Text = pThis.Source
    Set pThis.Model = viewModel
    Show
    I_CopyFetchFormView_ShowDialog = Not pThis.IsCancelled
End Function

Private Property Set I_CopyFetchFormView_Model(ByVal CFC As CopyFetchClass)
    Set pThis.Model = CFC
End Property

Private Property Get I_CopyFetchFormView_Model() As CopyFetchClass
    Set I_CopyFetchFormView_Model = pThis.Model
End Property

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub OtherDestinationButton_Click()
    pThis.Model.OtherDestinationClicked = True
End Sub

Private Sub OtherSourceButton_Click()
    pThis.Model.OtherSourceClicked = True
End Sub

Private Sub SourceBox_Change()
    SourceBox.Text = pThis.Source
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    pThis.IsCancelled = True
    Hide
    
    DestinationBox = vbNullString
    SourceBox = vbNullString
End Sub



