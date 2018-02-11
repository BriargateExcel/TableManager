VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyFetchClass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TView
    CopyClicked As Boolean
    OtherDestinationClicked As Boolean
    Destination As String
    
    FetchClicked As Boolean
    OtherSourceClicked As Boolean
    Source As String
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

Public Property Get OtherDestinationClicked() As Boolean
    OtherDestinationClicked = this.OtherDestinationClicked
End Property

Public Property Let OtherDestinationClicked(ByVal Clicked As Boolean)
    this.OtherDestinationClicked = Clicked
End Property

Public Property Get OtherSourceClicked() As Boolean
    OtherSourceClicked = this.OtherSourceClicked
End Property

Public Property Let OtherSourceClicked(ByVal Clicked As Boolean)
    this.OtherSourceClicked = Clicked
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal Src As String)
    this.Source = Src
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property

Public Property Let Destination(ByVal Dst As String)
    this.Destination = Dst
End Property


