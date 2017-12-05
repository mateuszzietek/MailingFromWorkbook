Attribute VB_Name = "FA_MailCreation"
Sub MessageFA()

Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object

Dim SigString As String
Dim Signature As String

Dim MessageText As String
MessageText = "<font face=Arial>Hi Team, <br> <br> Please process attached invoice. <u><b>Please inform us when the invoice gets <font color=#3399ff>booked</font> on Your side.</b></u><br> <br></font>"

'COPY TO TEMP SHEET & AUTOFIT
Sheets("FA").Range("A2:J300").SpecialCells(xlCellTypeConstants).Copy Destination:=Sheets("temp").Range("A1")
Sheets("temp").Range("A1:J300").Columns.AutoFit

Dim ReqNum
ReqNum = Sheets("temp").Range("A2").Value

Dim InvAttachment As String
InvAttachment = Sheets("temp").Range("C2").Value


'SIGNATURE
SigString = Environ("appdata") & "\Microsoft\Signatures\EUMarketing.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If


'SEND SEPARATELY?
Set rng = Nothing

Answer = MsgBox("Send ech one separately?", vbYesNo)
If Answer = vbYes Then
    Set rng = Sheets("temp").Range("A1:K2")
Else
    Set rng = Sheets("temp").Range("A1:K300").SpecialCells(xlCellTypeConstants)
End If

'EMAIL
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

On Error Resume Next

With OutMail
    .To = "APInvoiceNLHQstaples@Staples.com; supportoffice-ap@staples.com"
    .CC = "EUMarketingP2P@Staples-Solutions.com"
    .BCC = ""
    .Subject = "Processed Marketing Invoice " + Format(Now, "dd/mm/yyyy") + " (" + ReqNum + ")"
    .HTMLBody = MessageText + RangetoHTML(rng) + "<br> <br>" + Signature
    .SentOnBehalfOfName = "EUMarketingP2P@Staples-Solutions.com"
    .Attachments.Add ("G:\PTP Marketing\01. Operations\03. Europe Marketing Invoices\" + InvAttachment + ".pdf")
    .Display
    
End With

On Error GoTo 0

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set OutMail = Nothing
Set OutApp = Nothing

'CLEAR TEMP SHEET
Sheets("temp").Range("A:Q").Delete

End Sub


