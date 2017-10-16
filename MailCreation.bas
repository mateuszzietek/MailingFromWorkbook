Attribute VB_Name = "MailCreation"
Sub Mail_Selection_Range_Outlook_Body()
Attribute Mail_Selection_Range_Outlook_Body.VB_ProcData.VB_Invoke_Func = "S\n14"

Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object

Dim MessageText As String
MessageText = "Hi Team, <br> <br> Please process attached invoice. Please inform us when the invoice <b>gets booked</b> on Your side.<br> <br>"

Dim BU As String
BU = ActiveSheet.Range("I2").Value

Dim POAttachment As String
POAttachment = ActiveSheet.Range("B2").Value

'Skopiowanie danych do arkusza temp
Sheets("Sheet1").Range("A1:K300").SpecialCells(xlCellTypeConstants).Copy Destination:=Sheets("temp").Range("A1")

Sheets("temp").Range("A1:K300").Columns.AutoFit

'Czy wysylac pojedynczo?
Set rng = Nothing

Answer = MsgBox("Send ech one separately?", vbYesNo)
If Answer = vbYes Then
    Set rng = Sheets("temp").Range("A1:K2")
Else
    Set rng = Sheets("temp").Range("A1:K300").SpecialCells(xlCellTypeConstants)
End If

'Przygotowanie maila.
If rng Is Nothing Then
    MsgBox "The selection is not a range or the sheet is protected. " & _
           vbNewLine & "Please correct and try again.", vbOKOnly
    Exit Sub
End If

With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)


With OutMail
    .To = ""
    .CC = "EUMarketingP2P@Staples-Solutions.com; Przemyslaw.Luczak@Staples-Solutions.com"
    .BCC = ""
    .Subject = "AP Marketing Invoice " + BU
    .HTMLBody = MessageText + RangetoHTML(rng)
    .SentOnBehalfOfName = "EUMarketingP2P@Staples-Solutions.com"
    '.Attachments.Add ("G:\PTP Marketing\01. Operations\05. Finalised PO Folder FY 2017\" + POAttachment + ".pdf")
    
    ' In place of the following statement, you can use ".Display" to
    ' display the e-mail message.
    .Display
End With
On Error GoTo 0

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set OutMail = Nothing
Set OutApp = Nothing

'wyczyszczenie arkusza temp
Sheets("temp").Range("A1:Q1000").Clear

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
    RangetoHTML = ts.ReadAll
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



