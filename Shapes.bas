Attribute VB_Name = "Shapes"
Option Explicit

'========================================
'Name: DeleteAllShapes
'Purpose: Deletes all shapes from a sheet
'========================================

Sub DeleteAllShapes(wsSheets As Worksheet)
Dim wShape As Shape

For Each wShape In wsSheets.Shapes
    wShape.Delete
Next wShape
End Sub
