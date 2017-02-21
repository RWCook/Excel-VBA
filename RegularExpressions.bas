Attribute VB_Name = "RegularExpressions"
Option Explicit

'==================================================
'=Name   : RegexReplace
'=Purpose: To enable the use of regular expressions to replace
'character strings. Intended for use in a worksheet but MyRange
'can easily be changed to a string, for use in VBA.
'Usage:        RegexReplace(Range,PatternToMatch,ReplacementCharacters,Global)
'Example:    RegexReplace(A2,"^h","C",FALSE)
'==================================================
Function RegexReplace(MyRange As Range, _
            strMatch As String, _
            strReplace As String, _
            booGlobal As Boolean) As String
     
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
            .global = booGlobal
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
    End With

    RegexReplace = regex.Replace(MyRange, strReplace)

End Function

'==================================================
'=Name   : RegexMatch
'=Purpose: To enable the use of regular expressions to identify matched
'character strings. Intended for use in a worksheet but MyRange
'can easily be changed to a string, for use in VBA.
'Usage:        RegexMatch(Range,PatternToMatch)
'Example:    RegexMatch(A2,"^h")
'==================================================
Function RegexMatch(MyRange As Range, _
            strMatch As String _
            ) As Boolean
     
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

        With regex
            .global = False
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
        End With

If regex.Test(MyRange) Then
            RegexMatch = True
        Else: RegexMatch = False
End If

End Function

'==================================================
'=Name   : RegexMatchCount
'=Purpose: To count the number of matches based on a regular expression.
'Intended for use in a worksheet but MyRange
'can easily be changed to a string, for use in VBA.
'Usage:        RegexMatchCount(Range,PatternToMatch)
'Example:    RegexMatchCount(A2,"h")
'==================================================
Function RegexMatchCount(MyRange As Range, _
            strMatch As String) As Integer
     
Dim regex As Object
Dim RegexMatches As Object
Set regex = CreateObject("VBScript.RegExp")

        With regex
            .global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
        End With

Set RegexMatches = regex.Execute(MyRange)

RegexMatchCount = RegexMatches.Count()
           
End Function
