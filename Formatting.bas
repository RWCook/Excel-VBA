Attribute VB_Name = "Module1"
Option Explicit

'==================================================
'Name: FormatSheet
'Purpose: Very Basic formatting to set the heading
'and autofit the columns. Intended to be used with
'ApplyBorders to provide basic formatting.
'==================================================
Sub FormatSheet(wsSheet As Worksheet)

With wsSheet.Range("a1", Cells(1, wsSheet.UsedRange.Columns.Count))
    .Interior.ColorIndex = 35
    .Font.Name = "Verdana"
    .Font.Bold = True
    .Font.Italic = True
End With

wsSheet.UsedRange.Columns.AutoFit

End Sub

'==================================================
'Name:ApplyBorders
'Purpose: Very basic application of borders on the used 
'range.
'==================================================
Sub ApplyBorders(wsSheet As Worksheet)

With wsSheet.UsedRange.Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With


End Sub

'==================================================
'Name:ColourChart
'Purpose: Puts the 56 colours of Excel's ColorIndex
'on screen along with their numbers. Provides a quick
'way to look up the colours.
'==================================================
Sub ColourChart()
Dim intColumnNo As Integer
Dim intRowNo As Integer

intColumnNo = 1
intRowNo = 1

Dim i As Integer

For i = 1 To 56

    If i Mod 10 = 1 And intRowNo > 1 Then
        intColumnNo = intColumnNo + 2
        intRowNo = 1
    End If
    
    Cells(intRowNo, intColumnNo) = i
    Cells(intRowNo, intColumnNo + 1).Interior.ColorIndex = i
    
    intRowNo = intRowNo + 1
Next i
End Sub

'==================================================
'Name: RGBColours
'Purpose: Only really useful as a reminder of the RGB
'colour syntax.
'==================================================
Sub RGBColours()
Dim intRow As Integer
Dim intCol As Integer

For intRow = 1 To 25
    For intCol = 1 To 25
        Cells(intRow, intCol).Interior.Color = RGB(intRow * 10, intCol * 10, intRow * 10)
        
    Next intCol
Next intRow
    
End Sub
