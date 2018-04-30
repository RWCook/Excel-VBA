Attribute VB_Name = "Web"
Option Explicit

'=================================================
' Name : CallWebQuery
' Purpose: Example of how to call WebQueryFunction
'=================================================

Sub CallWebQuery()
Dim wbBook As Workbook
Dim wsSheet As Worksheet
Dim booOk As Boolean

Set wbBook = Application.ThisWorkbook
Set wsSheet = wbBook.Sheets(1)

Dim rngDestinationRange As Range
Set rngDestinationRange = wsSheet.Range(wsSheet.Cells(1, 1), wsSheet.Cells(1, 1))


Dim strUrl As String
strUrl = "https://en.wikipedia.org/wiki/Unicorn_(Tyrannosaurus_Rex_album)"
Dim strQueryTablename As String
strQueryTablename = "Tab1"
Dim strWebTables As String
strWebTables = "1,4,5" 'i.e. retrieve tables 1 and 4 and 5

booOk = WebQuery(wsSheet, True, rngDestinationRange, strUrl, strQueryTablename, strWebTables)

MsgBox prompt:="The query ran ok = " & booOk
End Sub

'==============================
'Name       : WebQuery
'Purpose    : Gets data from tables on the web and stores it in a worksheet.
'Parameters : wsDestination - the sheet where the data is going
'             booClearContents - if true then clear the used range of the destination sheet
'             rngDestinationRange - where to put the data on the destination sheet
'             strURL - the URL of the source
'             strQueryTableName - the name you wish to give the query table
'             strWebTables - the number or numbers of the web tables you want to retrieve
'Returns    : Returns true unless there is an error, in which case it returns false.
'==============================

Private Function WebQuery( _
        ByVal wsDestination As Worksheet, _
        ByVal booClearContents As Boolean, _
        ByVal rngDestinationRange As Range, _
        ByVal strUrl As String, _
        ByVal strQueryTablename As String, _
        ByVal strWebTables As String _
        ) As Boolean

Dim qt As QueryTable
On Error GoTo errHandler
wsDestination.Visible = xlSheetVisible
wsDestination.Select

If booClearContents = True Then
    wsDestination.UsedRange.ClearContents
    wsDestination.UsedRange.ClearFormats
End If

Set qt = wsDestination.QueryTables.Add(Connection:="URL;" & strUrl, Destination:=rngDestinationRange)
qt.RefreshOnFileOpen = False
qt.Name = strQueryTablename
qt.FieldNames = True
qt.WebSelectionType = xlSpecifiedTables
qt.WebTables = strWebTables
qt.Refresh BackgroundQuery:=False

WebQuery = True
Exit Function

errHandler:
    WebQuery = False
    If InStr(1, Err.Description, "Cannot locate the Internet server or proxy", vbTextCompare) > 0 Then
    MsgBox prompt:="Error getting data from the internet. Please check your connection. Macro stopped"
    Else:
    Debug.Print "Unknown Error in WebQuery" & Err.Number & ": " & Err.Description
    End If
    Exit Function
    
End Function
