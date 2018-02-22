VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyFetchClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TView
    CopyClicked As Boolean
    FetchClicked As Boolean
    ChangeFileClicked As Boolean
    FileName As String
End Type

Private this As TView

Public Property Get Self()
    Set Self = Me
End Property

Public Property Get CopyClicked() As Boolean
    CopyClicked = this.CopyClicked
End Property

Public Property Let CopyClicked(ByVal Clicked As Boolean)
    this.CopyClicked = Clicked
End Property

Public Property Get FetchClicked() As Boolean
    FetchClicked = this.FetchClicked
End Property

Public Property Let FetchClicked(ByVal Clicked As Boolean)
    this.FetchClicked = Clicked
End Property

Public Property Get ChangeFileClicked() As Boolean
    ChangeFileClicked = this.ChangeFileClicked
End Property

Public Property Let ChangeFileClicked(ByVal Clicked As Boolean)
    this.ChangeFileClicked = Clicked
End Property


