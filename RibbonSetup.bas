Attribute VB_Name = "RibbonSetup"
'@Folder("TableManager.Main")

Option Explicit

'**************************************************************
'AUTHOR: Chris Newman, TheSpreadsheetGuru
'Instructions on how to use this template can be found at:
'www.TheSpreadsheetGuru.com/blog/create-first-excel-addin
'**************************************************************

Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
'PURPOSE: Show/Hide buttons based on how many you need (False = Hide/True = Show)

Select Case control.ID
  Case "GroupA": MakeVisible = True
  Case "aButton01": MakeVisible = True
  Case "aButton02": MakeVisible = True
  Case "aButton03": MakeVisible = True
  Case "aButton04": MakeVisible = False
  Case "aButton05": MakeVisible = False
  Case "aButton06": MakeVisible = False
  Case "aButton07": MakeVisible = False
  Case "aButton08": MakeVisible = False
  Case "aButton09": MakeVisible = False
  Case "aButton10": MakeVisible = False
  
  Case "GroupB": MakeVisible = False
  Case "bButton01": MakeVisible = True
  Case "bButton02": MakeVisible = True
  Case "bButton03": MakeVisible = True
  Case "bButton04": MakeVisible = True
  Case "bButton05": MakeVisible = True
  Case "bButton06": MakeVisible = True
  Case "bButton07": MakeVisible = True
  Case "bButton08": MakeVisible = True
  Case "bButton09": MakeVisible = True
  Case "bButton10": MakeVisible = True
  
  Case "GroupC": MakeVisible = False
  Case "cButton01": MakeVisible = True
  Case "cButton02": MakeVisible = True
  Case "cButton03": MakeVisible = True
  Case "cButton04": MakeVisible = True
  Case "cButton05": MakeVisible = True
  Case "cButton06": MakeVisible = True
  Case "cButton07": MakeVisible = True
  Case "cButton08": MakeVisible = True
  Case "cButton09": MakeVisible = True
  Case "cButton10": MakeVisible = True
  
  Case "GroupD": MakeVisible = False
  Case "dButton01": MakeVisible = True
  Case "dButton02": MakeVisible = True
  Case "dButton03": MakeVisible = True
  Case "dButton04": MakeVisible = True
  Case "dButton05": MakeVisible = True
  Case "dButton06": MakeVisible = True
  Case "dButton07": MakeVisible = True
  Case "dButton08": MakeVisible = True
  Case "dButton09": MakeVisible = True
  Case "dButton10": MakeVisible = True
  
  Case "GroupE": MakeVisible = False
  Case "eButton01": MakeVisible = True
  Case "eButton02": MakeVisible = True
  Case "eButton03": MakeVisible = True
  Case "eButton04": MakeVisible = True
  Case "eButton05": MakeVisible = True
  Case "eButton06": MakeVisible = True
  Case "eButton07": MakeVisible = True
  Case "eButton08": MakeVisible = True
  Case "eButton09": MakeVisible = True
  Case "eButton10": MakeVisible = True
  
  Case "GroupF": MakeVisible = False
  Case "fButton01": MakeVisible = True
  Case "fButton02": MakeVisible = True
  Case "fButton03": MakeVisible = True
  Case "fButton04": MakeVisible = True
  Case "fButton05": MakeVisible = True
  Case "fButton06": MakeVisible = True
  Case "fButton07": MakeVisible = True
  Case "fButton08": MakeVisible = True
  Case "fButton09": MakeVisible = True
  Case "fButton10": MakeVisible = True
  
End Select

End Sub

Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling)
'PURPOSE: Determine the text to go along with your Tab, Groups, and Buttons

Select Case control.ID
  
  Case "CustomTab": Labeling = "Table Manager"
  
  Case "GroupA": Labeling = "Overall"
  Case "aButton01": Labeling = "Build Forms"
  Case "aButton02": Labeling = "Build Data Descriptions"
  Case "aButton03": Labeling = "Extend Data Validation"
  Case "aButton04": Labeling = "Button"
  Case "aButton05": Labeling = "Button"
  Case "aButton06": Labeling = "Button"
  Case "aButton07": Labeling = "Button"
  Case "aButton08": Labeling = "Button"
  Case "aButton09": Labeling = "Button"
  Case "aButton10": Labeling = "Button"
  
  Case "GroupB": Labeling = "Group Name"
  Case "bButton01": Labeling = "Button"
  Case "bButton02": Labeling = "Button"
  Case "bButton03": Labeling = "Button"
  Case "bButton04": Labeling = "Button"
  Case "bButton05": Labeling = "Button"
  Case "bButton06": Labeling = "Button"
  Case "bButton07": Labeling = "Button"
  Case "bButton08": Labeling = "Button"
  Case "bButton09": Labeling = "Button"
  Case "bButton10": Labeling = "Button"
  
  Case "GroupC": Labeling = "Group Name"
  Case "cButton01": Labeling = "Button"
  Case "cButton02": Labeling = "Button"
  Case "cButton03": Labeling = "Button"
  Case "cButton04": Labeling = "Button"
  Case "cButton05": Labeling = "Button"
  Case "cButton06": Labeling = "Button"
  Case "cButton07": Labeling = "Button"
  Case "cButton08": Labeling = "Button"
  Case "cButton09": Labeling = "Button"
  Case "cButton10": Labeling = "Button"
  
  Case "GroupD": Labeling = "Group Name"
  Case "dButton01": Labeling = "Button"
  Case "dButton02": Labeling = "Button"
  Case "dButton03": Labeling = "Button"
  Case "dButton04": Labeling = "Button"
  Case "dButton05": Labeling = "Button"
  Case "dButton06": Labeling = "Button"
  Case "dButton07": Labeling = "Button"
  Case "dButton08": Labeling = "Button"
  Case "dButton09": Labeling = "Button"
  Case "dButton10": Labeling = "Button"
  
  Case "GroupE": Labeling = "Group Name"
  Case "eButton01": Labeling = "Button"
  Case "eButton02": Labeling = "Button"
  Case "eButton03": Labeling = "Button"
  Case "eButton04": Labeling = "Button"
  Case "eButton05": Labeling = "Button"
  Case "eButton06": Labeling = "Button"
  Case "eButton07": Labeling = "Button"
  Case "eButton08": Labeling = "Button"
  Case "eButton09": Labeling = "Button"
  Case "eButton10": Labeling = "Button"
  
  Case "GroupF": Labeling = "Group Name"
  Case "fButton01": Labeling = "Button"
  Case "fButton02": Labeling = "Button"
  Case "fButton03": Labeling = "Button"
  Case "fButton04": Labeling = "Button"
  Case "fButton05": Labeling = "Button"
  Case "fButton06": Labeling = "Button"
  Case "fButton07": Labeling = "Button"
  Case "fButton08": Labeling = "Button"
  Case "fButton09": Labeling = "Button"
  Case "fButton10": Labeling = "Button"
  
End Select
   
End Sub

Sub GetImage(control As IRibbonControl, ByRef RibbonImage)
'PURPOSE: Tell each button which image to load from the Microsoft Icon Library
'TIPS: Image names are case sensitive, if image does not appear in ribbon after re-starting Excel, the image name is incorrect

Select Case control.ID
  
  Case "aButton01": RibbonImage = "ObjectPictureFill"
  Case "aButton02": RibbonImage = "ObjectPictureFill"
  Case "aButton03": RibbonImage = "ObjectPictureFill"
  Case "aButton04": RibbonImage = "ObjectPictureFill"
  Case "aButton05": RibbonImage = "ObjectPictureFill"
  Case "aButton06": RibbonImage = "ObjectPictureFill"
  Case "aButton07": RibbonImage = "ObjectPictureFill"
  Case "aButton08": RibbonImage = "ObjectPictureFill"
  Case "aButton09": RibbonImage = "ObjectPictureFill"
  Case "aButton10": RibbonImage = "ObjectPictureFill"
  
  Case "bButton01": RibbonImage = "ObjectPictureFill"
  Case "bButton02": RibbonImage = "ObjectPictureFill"
  Case "bButton03": RibbonImage = "ObjectPictureFill"
  Case "bButton04": RibbonImage = "ObjectPictureFill"
  Case "bButton05": RibbonImage = "ObjectPictureFill"
  Case "bButton06": RibbonImage = "ObjectPictureFill"
  Case "bButton07": RibbonImage = "ObjectPictureFill"
  Case "bButton08": RibbonImage = "ObjectPictureFill"
  Case "bButton09": RibbonImage = "ObjectPictureFill"
  Case "bButton10": RibbonImage = "ObjectPictureFill"
  
  Case "cButton01": RibbonImage = "ObjectPictureFill"
  Case "cButton02": RibbonImage = "ObjectPictureFill"
  Case "cButton03": RibbonImage = "ObjectPictureFill"
  Case "cButton04": RibbonImage = "ObjectPictureFill"
  Case "cButton05": RibbonImage = "ObjectPictureFill"
  Case "cButton06": RibbonImage = "ObjectPictureFill"
  Case "cButton07": RibbonImage = "ObjectPictureFill"
  Case "cButton08": RibbonImage = "ObjectPictureFill"
  Case "cButton09": RibbonImage = "ObjectPictureFill"
  Case "cButton10": RibbonImage = "ObjectPictureFill"
  
  Case "dButton01": RibbonImage = "ObjectPictureFill"
  Case "dButton02": RibbonImage = "ObjectPictureFill"
  Case "dButton03": RibbonImage = "ObjectPictureFill"
  Case "dButton04": RibbonImage = "ObjectPictureFill"
  Case "dButton05": RibbonImage = "ObjectPictureFill"
  Case "dButton06": RibbonImage = "ObjectPictureFill"
  Case "dButton07": RibbonImage = "ObjectPictureFill"
  Case "dButton08": RibbonImage = "ObjectPictureFill"
  Case "dButton09": RibbonImage = "ObjectPictureFill"
  Case "dButton10": RibbonImage = "ObjectPictureFill"
  
  Case "eButton01": RibbonImage = "ObjectPictureFill"
  Case "eButton02": RibbonImage = "ObjectPictureFill"
  Case "eButton03": RibbonImage = "ObjectPictureFill"
  Case "eButton04": RibbonImage = "ObjectPictureFill"
  Case "eButton05": RibbonImage = "ObjectPictureFill"
  Case "eButton06": RibbonImage = "ObjectPictureFill"
  Case "eButton07": RibbonImage = "ObjectPictureFill"
  Case "eButton08": RibbonImage = "ObjectPictureFill"
  Case "eButton09": RibbonImage = "ObjectPictureFill"
  Case "eButton10": RibbonImage = "ObjectPictureFill"
  
  Case "fButton01": RibbonImage = "ObjectPictureFill"
  Case "fButton02": RibbonImage = "ObjectPictureFill"
  Case "fButton03": RibbonImage = "ObjectPictureFill"
  Case "fButton04": RibbonImage = "ObjectPictureFill"
  Case "fButton05": RibbonImage = "ObjectPictureFill"
  Case "fButton06": RibbonImage = "ObjectPictureFill"
  Case "fButton07": RibbonImage = "ObjectPictureFill"
  Case "fButton08": RibbonImage = "ObjectPictureFill"
  Case "fButton09": RibbonImage = "ObjectPictureFill"
  Case "fButton10": RibbonImage = "ObjectPictureFill"
  
End Select

End Sub

Sub GetSize(control As IRibbonControl, ByRef Size)
'PURPOSE: Determine if the button size is large or small

Const Large As Integer = 1
Const Small As Integer = 0

Select Case control.ID
    
  Case "aButton01": Size = Large
  Case "aButton02": Size = Small
  Case "aButton03": Size = Small
  Case "aButton04": Size = Small
  Case "aButton05": Size = Small
  Case "aButton06": Size = Small
  Case "aButton07": Size = Small
  Case "aButton08": Size = Small
  Case "aButton09": Size = Small
  Case "aButton10": Size = Small
  
  Case "bButton01": Size = Large
  Case "bButton02": Size = Small
  Case "bButton03": Size = Small
  Case "bButton04": Size = Small
  Case "bButton05": Size = Small
  Case "bButton06": Size = Small
  Case "bButton07": Size = Small
  Case "bButton08": Size = Small
  Case "bButton09": Size = Small
  Case "bButton10": Size = Small
  
  Case "cButton01": Size = Large
  Case "cButton02": Size = Small
  Case "cButton03": Size = Small
  Case "cButton04": Size = Small
  Case "cButton05": Size = Small
  Case "cButton06": Size = Small
  Case "cButton07": Size = Small
  Case "cButton08": Size = Small
  Case "cButton09": Size = Small
  Case "cButton10": Size = Small
  
  Case "dButton01": Size = Large
  Case "dButton02": Size = Small
  Case "dButton03": Size = Small
  Case "dButton04": Size = Small
  Case "dButton05": Size = Small
  Case "dButton06": Size = Small
  Case "dButton07": Size = Small
  Case "dButton08": Size = Small
  Case "dButton09": Size = Small
  Case "dButton10": Size = Small
  
  Case "eButton01": Size = Large
  Case "eButton02": Size = Small
  Case "eButton03": Size = Small
  Case "eButton04": Size = Small
  Case "eButton05": Size = Small
  Case "eButton06": Size = Small
  Case "eButton07": Size = Small
  Case "eButton08": Size = Small
  Case "eButton09": Size = Small
  Case "eButton10": Size = Small
  
  Case "fButton01": Size = Large
  Case "fButton02": Size = Small
  Case "fButton03": Size = Small
  Case "fButton04": Size = Small
  Case "fButton05": Size = Small
  Case "fButton06": Size = Small
  Case "fButton07": Size = Small
  Case "fButton08": Size = Small
  Case "fButton09": Size = Small
  Case "fButton10": Size = Small
  
End Select

End Sub

Sub RunMacro(control As IRibbonControl)
'PURPOSE: Tell each button which macro subroutine to run when clicked

Select Case control.ID
  
  Case "aButton01": Application.Run "Auto_Open"
  Case "aButton02": Application.Run "BuildDataDescriptionTable"
  Case "aButton03": Application.Run "ExtendDataValidation"
  Case "aButton04": Application.Run "DummyMacro"
  Case "aButton05": Application.Run "DummyMacro"
  Case "aButton06": Application.Run "DummyMacro"
  Case "aButton07": Application.Run "DummyMacro"
  Case "aButton08": Application.Run "DummyMacro"
  Case "aButton09": Application.Run "DummyMacro"
  Case "aButton10": Application.Run "DummyMacro"
  
  Case "bButton01": Application.Run "DummyMacro"
  Case "bButton02": Application.Run "DummyMacro"
  Case "bButton03": Application.Run "DummyMacro"
  Case "bButton04": Application.Run "DummyMacro"
  Case "bButton05": Application.Run "DummyMacro"
  Case "bButton06": Application.Run "DummyMacro"
  Case "bButton07": Application.Run "DummyMacro"
  Case "bButton08": Application.Run "DummyMacro"
  Case "bButton09": Application.Run "DummyMacro"
  Case "bButton10": Application.Run "DummyMacro"
  
  Case "cButton01": Application.Run "DummyMacro"
  Case "cButton02": Application.Run "DummyMacro"
  Case "cButton03": Application.Run "DummyMacro"
  Case "cButton04": Application.Run "DummyMacro"
  Case "cButton05": Application.Run "DummyMacro"
  Case "cButton06": Application.Run "DummyMacro"
  Case "cButton07": Application.Run "DummyMacro"
  Case "cButton08": Application.Run "DummyMacro"
  Case "cButton09": Application.Run "DummyMacro"
  Case "cButton10": Application.Run "DummyMacro"
  
  Case "dButton01": Application.Run "DummyMacro"
  Case "dButton02": Application.Run "DummyMacro"
  Case "dButton03": Application.Run "DummyMacro"
  Case "dButton04": Application.Run "DummyMacro"
  Case "dButton05": Application.Run "DummyMacro"
  Case "dButton06": Application.Run "DummyMacro"
  Case "dButton07": Application.Run "DummyMacro"
  Case "dButton08": Application.Run "DummyMacro"
  Case "dButton09": Application.Run "DummyMacro"
  Case "dButton10": Application.Run "DummyMacro"
  
  Case "eButton01": Application.Run "DummyMacro"
  Case "eButton02": Application.Run "DummyMacro"
  Case "eButton03": Application.Run "DummyMacro"
  Case "eButton04": Application.Run "DummyMacro"
  Case "eButton05": Application.Run "DummyMacro"
  Case "eButton06": Application.Run "DummyMacro"
  Case "eButton07": Application.Run "DummyMacro"
  Case "eButton08": Application.Run "DummyMacro"
  Case "eButton09": Application.Run "DummyMacro"
  Case "eButton10": Application.Run "DummyMacro"
  
  Case "fButton01": Application.Run "DummyMacro"
  Case "fButton02": Application.Run "DummyMacro"
  Case "fButton03": Application.Run "DummyMacro"
  Case "fButton04": Application.Run "DummyMacro"
  Case "fButton05": Application.Run "DummyMacro"
  Case "fButton06": Application.Run "DummyMacro"
  Case "fButton07": Application.Run "DummyMacro"
  Case "fButton08": Application.Run "DummyMacro"
  Case "fButton09": Application.Run "DummyMacro"
  Case "fButton10": Application.Run "DummyMacro"

 End Select
    
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef Screentip)
'PURPOSE: Display a specific macro description when the mouse hovers over a button

Select Case control.ID
  
  Case "aButton01": Screentip = "Build a form for each table"
  Case "aButton02": Screentip = "Build the Data Description Table"
  Case "aButton03": Screentip = "Extend Data Validation down all rows of all tables"
  Case "aButton04": Screentip = "Description"
  Case "aButton05": Screentip = "Description"
  Case "aButton06": Screentip = "Description"
  Case "aButton07": Screentip = "Description"
  Case "aButton08": Screentip = "Description"
  Case "aButton09": Screentip = "Description"
  Case "aButton10": Screentip = "Description"
  
  Case "bButton01": Screentip = "Description"
  Case "bButton02": Screentip = "Description"
  Case "bButton03": Screentip = "Description"
  Case "bButton04": Screentip = "Description"
  Case "bButton05": Screentip = "Description"
  Case "bButton06": Screentip = "Description"
  Case "bButton07": Screentip = "Description"
  Case "bButton08": Screentip = "Description"
  Case "bButton09": Screentip = "Description"
  Case "bButton10": Screentip = "Description"
  
  Case "cButton01": Screentip = "Description"
  Case "cButton02": Screentip = "Description"
  Case "cButton03": Screentip = "Description"
  Case "cButton04": Screentip = "Description"
  Case "cButton05": Screentip = "Description"
  Case "cButton06": Screentip = "Description"
  Case "cButton07": Screentip = "Description"
  Case "cButton08": Screentip = "Description"
  Case "cButton09": Screentip = "Description"
  Case "cButton10": Screentip = "Description"
  
  Case "dButton01": Screentip = "Description"
  Case "dButton02": Screentip = "Description"
  Case "dButton03": Screentip = "Description"
  Case "dButton04": Screentip = "Description"
  Case "dButton05": Screentip = "Description"
  Case "dButton06": Screentip = "Description"
  Case "dButton07": Screentip = "Description"
  Case "dButton08": Screentip = "Description"
  Case "dButton09": Screentip = "Description"
  Case "dButton10": Screentip = "Description"

  Case "eButton01": Screentip = "Description"
  Case "eButton02": Screentip = "Description"
  Case "eButton03": Screentip = "Description"
  Case "eButton04": Screentip = "Description"
  Case "eButton05": Screentip = "Description"
  Case "eButton06": Screentip = "Description"
  Case "eButton07": Screentip = "Description"
  Case "eButton08": Screentip = "Description"
  Case "eButton09": Screentip = "Description"
  Case "eButton10": Screentip = "Description"
  
  Case "fButton01": Screentip = "Description"
  Case "fButton02": Screentip = "Description"
  Case "fButton03": Screentip = "Description"
  Case "fButton04": Screentip = "Description"
  Case "fButton05": Screentip = "Description"
  Case "fButton06": Screentip = "Description"
  Case "fButton07": Screentip = "Description"
  Case "fButton08": Screentip = "Description"
  Case "fButton09": Screentip = "Description"
  Case "fButton10": Screentip = "Description"
  
End Select

End Sub
