Attribute VB_Name = "Module1"
Option Explicit
'==================================================
'Name: CallComment
'Purpose: Sample calls for subroutines
'==================================================
Sub CallComment()
Dim wsSheet As Worksheet
Set wsSheet = ActiveSheet

Call DeleteComment(wsSheet.Range("A1"))

Call AddComment(wsSheet.Range("A1"), "Here is a comment")

Call ResizeComment(wsSheet.Range("a1"), 1, 9)

Call FormatCommentFont(wsSheet.Range("a1"), 3, 14, "Arial")

Call SetCommentBackgroundColour(wsSheet.Range("a1"), 16)

End Sub
'==================================================
'Name: ResizeComment
'Purpose: ResizeComment
'==================================================
Sub ResizeComment(MyRange As Range, dblHeight As Double, dblWidth As Double)
MyRange.comment.Shape.ScaleWidth dblWidth, msoFalse, msoScaleFromTopLeft
MyRange.comment.Shape.ScaleHeight dblHeight, msoFalse, msoScaleFromTopLeft

End Sub
'==================================================
'Name: DeleteComment
'Purpose: Delete a comment from a cell
'==================================================
Sub DeleteComment(MyRange As Range)
MyRange.comment.Delete
End Sub
'==================================================
'Name: AddComment
'Purpose: To add comment text to a cell.
'==================================================
Sub AddComment(MyRange As Range, strCommentText As String)

MyRange.Select

MyRange.AddComment (strCommentText)
End Sub
'==================================================
'Name: SetCommentBackgroundColour
'Purpose: Set the background colour of a comment
'==================================================
Sub SetCommentBackgroundColour(MyRange As Range, intBackgroundColour As Integer)
MyRange.comment.Shape.Fill.ForeColor.SchemeColor = intBackgroundColour
End Sub
'==================================================
'Name: FormatCommentFont
'Purpose: To change the comments font, font size and font colour
'==================================================
Sub FormatCommentFont(MyRange As Range, intColour As Integer, intFontSize As Integer, strFontName As String)

With MyRange.comment.Shape.TextFrame.Characters.Font
    .ColorIndex = intColour
    .Size = intFontSize
    .Name = strFontName

End With

End Sub
'==================================================
'Name: RangeHasComment
'Purpose: To detect whether a range has a comment
'==================================================
Function RangeHasComment(MyRange As Range) As Boolean
 
If MyRange.comment Is Nothing Then
    RangeHasComment = False
Else
    RangeHasComment = True
End If

End Function
'==================================================
'Name:GetCommentText
'Purpose: To get the text from a comment and place it in a cell.
'Intended to be used on a worksheet.
'==================================================
Function GetCommentText(MyRange As Range) As String
 
If MyRange.comment Is Nothing Then
    GetCommentText = ""
Else
    GetCommentText = MyRange.comment.Text
End If

End Function
'==================================================
'Name: CommentTextMatch
'Purpose: To find text matches inside comments. Uses regular expressions
' through the RegexMatch function below to do this.
'Example: CommentTextMatch(A1,"^h") or CommentTextMatch(A1,"hat")
'==================================================
Function CommentTextMatch(MyRange As Range, myStringToMatch As String) As Boolean
Dim strText As String
Dim varRes As Variant

If MyRange.comment Is Nothing Then
    CommentTextMatch = False
Else
CommentTextMatch = RegexMatch(MyRange.comment.Text, myStringToMatch)
End If

End Function

'==================================================
'=Name   : RegexMatch
'=Purpose: To enable the use of regular expressions to identify matched
'character strings. In this context, for use by CommentTextMatch
'Usage:        RegexMatch(String,PatternToMatch)
'Example:    RegexMatch("hat","^h")
'==================================================
Function RegexMatch(MyString As String, _
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

If regex.Test(MyString) Then
            RegexMatch = True
        Else: RegexMatch = False
End If

End Function
'==================================================
'Name:    ListComments
'Purpose:Lists all comments in the ActiveSheet to the Immediate window
'==================================================
Sub ListComments()
Dim cmt As comment
For Each cmt In ActiveSheet.comments
    Debug.Print cmt.Parent.Address & ": " & cmt.Text
Next cmt

End Sub

