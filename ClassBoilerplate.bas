Attribute VB_Name = "ClassBoilerplate"
Option Explicit
Option Base 1

'***********************
'Name:    Make_Object_Boilerplate
'Purpose: To create boilerplate class code in the immediate window.
'This is intended to be used for very simple classes where the set up of
'a class seems like a lot of typing, but the clarity of using a class would
'be beneficial.
'Inputs:  The macro works on the 1st worksheet. Row 1 contains properties
'(i.e. the macro will read all of the headings and create properties for
'each of them). Row 2 needs to have the data types
'for each of these headings.

Public Sub Make_Object_Boilerplate()

Dim arrData() As Variant
Dim wsThisSheet As Worksheet
Set wsThisSheet = ThisWorkbook.Sheets(1)


ReDim arrData(wsThisSheet.UsedRange.Columns.Count, 2)

Dim i As Integer
For i = 1 To wsThisSheet.UsedRange.Columns.Count
    arrData(i, 1) = wsThisSheet.Cells(1, i)
    arrData(i, 1) = reformat(arrData(i, 1))
    
    If wsThisSheet.Cells(2, i) = vbNullString Then
        arrData(i, 2) = "Variant"
    Else
            arrData(i, 2) = wsThisSheet.Cells(2, i)
    End If
    
Next i

'Create Basic Class Code in Immediate Window
'Set Variables
    Debug.Print "Option explicit"
For i = 1 To UBound(arrData, 1)
    Debug.Print "Private p" & arrData(i, 1) & " as " & arrData(i, 2)
Next i

For i = 1 To UBound(arrData, 1)
    'Property Let
    Debug.Print "Property Let " & arrData(i, 1) & " (ByVal value as " & arrData(i, 2) & ")"
    Debug.Print "p" & arrData(i, 1) & "=value"
    Debug.Print "End property"
    
    'Property Get
    Debug.Print "Property Get " & arrData(i, 1) & "() as " & arrData(i, 2)
        Debug.Print arrData(i, 1) & "= p" & arrData(i, 1)
    Debug.Print "End property"
Next i

End Sub
Private Function reformat(ByVal value As String) As String
reformat = Replace(value, " ", "_")

End Function
