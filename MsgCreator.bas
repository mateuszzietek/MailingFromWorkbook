Attribute VB_Name = "MsgCreator"
Sub Message_Creator()

Dim rng As Range
Dim OutApp, OutMail As Object
Dim MessageText, SigString, Signature As String
Dim Recipient, Subject, YourMailbox As String

'FILL THIS DATA BEFORE OPERATING
Recipient = ""
Subject = "" & Date
YourMailbox = ""
YourOutlookSignatureName = ""

On Error GoTo NOTIFICATION:

'EMAIL TEXT
MessageText = _
"<font face=Arial>Hi, <br> <br> Please find the table. <u><b><font color=#3399ff> This part of text is formatedy by HTML tags.</b></u><br> <br></font>"

'EXPORT SELCTED RANGE TO TEMP SHEET
Selection.Copy Destination:=Sheets("temp").Range("A1")
Sheets("temp").Range("A1:Q100").Columns.AutoFit

'SIGNATURE
SigString = Environ("appdata") & "\Microsoft\Signatures\" & YourOutlookSignatureName & ".htm"
    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If

'RANGE FOR E-MAIL CONTENT
Set rng = Nothing
Set rng = Sheets("temp").Range("A1:Q500").SpecialCells(xlCellTypeVisible)
SingleEmail = 0

'EMAIL
With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

On Error Resume Next

With OutMail
    .To = Recipient
    .CC = ""
    .BCC = ""
    .Subject = Subject
    .HTMLBody = MessageText + RangetoHTML(rng) + "<br> <br>" + Signature
    .SentOnBehalfOfName = YourMailbox
    .Display
End With


'ATACHMENTS

'This part of code is designed for automatic attaching files from specified folder.

'provide column number (from selected range) with filenames data
ColumnNumber = 3

'provide folder path
FolderPath = "G:\folder\subfolder\"

'specify file type
FileType = ".pdf"

' the loop starts from 2 to because the first row is usually a header.
' "If statement" identifies duplicates and avoid to attached them.
' It's possible to use - OutMail.Attachments.Add - many times in different loops to create multiple attachments

For n = 2 To 999
    
    Filename = Sheets("temp").Cells(n, ColumnNumber).Value
    
    If Sheets("temp").Cells(n, ColumnNumber).Value = Sheets("temp").Cells(n - 1, ColumnNumber).Value Then
            FileToAttach = ""
    Else
            FileToAttach = FolderPath + Filename + FileType
    End If
            OutMail.Attachments.Add (FileToAttach)

Next n


'ENDING SUB
Sheets("temp").Range("A:X").Delete

With Application

    .EnableEvents = True
    .ScreenUpdating = True
    
End With

Set OutMail = Nothing
Set OutApp = Nothing

Exit Sub

NOTIFICATION:
    MsgBox ("Make sure that you have cretaed sheet named - temp")

End Sub


Function RangetoHTML(rng As Range)

' By Ron de Bruin.
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function


Function GetBoiler(ByVal sFile As String) As String

'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function





