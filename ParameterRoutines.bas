Attribute VB_Name = "ParameterRoutines"
Option Explicit

Public Function DarkestColorValue() As Long
    DarkestColorValue = TableManager.GetCellValue("ColorTable", "Color Name", "Darkest Color", "Decimal Color Value")
End Function

Public Function LightestColorValue() As Long
    LightestColorValue = TableManager.GetCellValue("ColorTable", "Color Name", "Lightest Color", "Decimal Color Value")
End Function


