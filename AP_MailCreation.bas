Attribute VB_Name = "AP_MailCreation"
Sub Mail_Selection_Range_Outlook_Body()
Attribute Mail_Selection_Range_Outlook_Body.VB_ProcData.VB_Invoke_Func = "S\n14"

Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object

Dim MessageText As String
MessageText = "Hi Team, <br> <br> Please process attached invoice. <u><b>Please inform us when the invoice gets booked on Your side.</b></u><br> <br>"


'Skopiowanie danych do arkusza temp
Sheets("AP").Range("A1:K300").SpecialCells(xlCellTypeConstants).Copy Destination:=Sheets("temp").Range("A1")

Sheets("temp").Range("A1:K300").Columns.AutoFit

Dim BU As String
BU = Sheets("temp").Range("I2").Value

Dim POAttachment As String
POAttachment = Sheets("temp").Range("B2").Value


'Przypisanie maila do BU
Dim BUmail As String

Select Case BU
    
        Case Is = "AT - SE Advantage Austria"
                    BUmail = Sheets("apc").Range("C26").Value

        Case Is = "AT - SE SBD Austria (not in use)"
                    BUmail = Sheets("apc").Range("C26").Value
        
        Case Is = "AT - SE SBD Pressel (Austria)"
                    BUmail = Sheets("apc").Range("C24").Value
        
        Case Is = "BE - SE Advantage Belgium"
                    BUmail = Sheets("apc").Range("C38").Value
        
        Case Is = "BE - SE Bernard Belgium"
                    BUmail = Sheets("apc").Range("C37").Value
        
        Case Is = "DE - SE Advantage Germany"
                    BUmail = Sheets("apc").Range("C18").Value
        
        Case Is = "DE - SE Retail Germany"
                    BUmail = Sheets("apc").Range("C17").Value
        
        Case Is = "DE - SE SBD Germany Total"
                    BUmail = Sheets("apc").Range("C19").Value
        
        Case Is = "DE - SE SBD Pressel (Germany)"
                    BUmail = Sheets("apc").Range("C23").Value
        
        Case Is = "DK - SE Advantage Denmark"
                    BUmail = Sheets("apc").Range("C5").Value
        
        Case Is = "DK - SE SBD Denmark"
                    BUmail = Sheets("apc").Range("C5").Value
        
        Case Is = "ES - SE Advantage Spain"
                    BUmail = Sheets("apc").Range("C28").Value
        
        Case Is = "ES - SE SBD Spain"
                    BUmail = Sheets("apc").Range("C28").Value
        
'        Case Is = "EU - Staples Europe BV"
'                    BUmail = Sheets("apc").Range("C00").Value
        
'        Case Is = "EU - Staples Europe Import BV"
'                    BUmail = Sheets("apc").Range("C00").Value
        
'        Case Is = "EU - Staples International BV"
'                    BUmail = Sheets("apc").Range("C00").Value
        
        Case Is = "FI - SE Advantage Finland"
                    BUmail = Sheets("apc").Range("C34").Value
        
        Case Is = "FI - SE Holding Finland"
                    BUmail = Sheets("apc").Range("C34").Value
        
        Case Is = "FR - SE Advantage France"
                    BUmail = Sheets("apc").Range("C13").Value
        
        Case Is = "FR - SE Bernard France"
                    BUmail = Sheets("apc").Range("C13").Value
        
        Case Is = "FR - SE SBD France"
                    BUmail = Sheets("apc").Range("C13").Value
        
        Case Is = "IT - SE Advantage Italy"
                    BUmail = Sheets("apc").Range("C30").Value
        
        Case Is = "IT - SE SBD Italy Total"
                    BUmail = Sheets("apc").Range("C30").Value
        
        Case Is = "NL - SE Advantage Netherlands Total"
                    BUmail = Sheets("apc").Range("C40").Value
        
        Case Is = "NL - SE Retail Netherlands"
                    BUmail = Sheets("apc").Range("C40").Value
        
        Case Is = "NL - SE SBD Netherlands"
                    BUmail = Sheets("apc").Range("C40").Value
        
        Case Is = "NO - EMO Norway"
                    BUmail = Sheets("apc").Range("C3").Value
        
        Case Is = "NO - SE Advantage Norway"
                    BUmail = Sheets("apc").Range("C3").Value
        
        Case Is = "NO - SE Retail Norway"
                    BUmail = Sheets("apc").Range("C3").Value
        
        Case Is = "NO - SE SBD Norway"
                    BUmail = Sheets("apc").Range("C3").Value
        
        Case Is = "PL - SE Advantage Poland"
                    BUmail = Sheets("apc").Range("C42").Value
        
        Case Is = "PT - SE Retail Portugal"
                    BUmail = "DO NOT SEND THIS EMAIL!"
        
        Case Is = "PT - SE SBD Portugal"
                    BUmail = "DO NOT SEND THIS EMAIL!"
        
        Case Is = "SE - EMO Sweden"
                    BUmail = Sheets("apc").Range("C4").Value
        
        Case Is = "SE - SE Advantage Sweden"
                    BUmail = Sheets("apc").Range("C4").Value
            
        Case Is = "SE - SE SBD Sweden"
                    BUmail = Sheets("apc").Range("C4").Value
        
        Case Is = "UK - SE Retail UK (not in use)"
                    BUmail = Sheets("apc").Range("C8").Value
        
        Case Is = "UK - Staples UK Adv Limited"
                    BUmail = Sheets("apc").Range("C9").Value
        
        Case Is = "UK - Staples UK Online Limited"
                    BUmail = Sheets("apc").Range("C10").Value

End Select

'Sygnatura
Dim SigString As String
Dim Signature As String
    
SigString = Environ("appdata") & "\Microsoft\Signatures\EUMarketing.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If


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
    .To = BUmail
    .CC = "EUMarketingP2P@Staples-Solutions.com; Przemyslaw.Luczak@Staples-Solutions.com"
    .BCC = ""
    .Subject = "AP Marketing Invoice " + BU
    .HTMLBody = MessageText + RangetoHTML(rng) + "<br> <br>" + Signature
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



