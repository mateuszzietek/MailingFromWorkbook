Attribute VB_Name = "AP_MailCreation"
Sub MessageAP()
Attribute MessageAP.VB_ProcData.VB_Invoke_Func = " \n14"

Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object

Dim MessageText As String

Dim BU As String
Dim BUmail As String

Dim SigString As String
Dim Signature As String

Dim InvoiceNumber As String
Dim InvoiceAttachment As String
Dim n As Integer
Dim rng2 As Range

Dim SingleEmail As Integer

Dim PO As String
Dim PoPdf As String
Dim PONumber As String
Dim PONumberDir As String

'Dim PONumber2018
'Dim PoPdf2018 As String
'Dim PONumberDir2018 As String

Dim m As Integer

Dim Msg As String
Dim ReqNumber As String
Dim MsgFile As String
Dim MsgAttachmentDir As String
Dim MsgAttachment As String

Dim InvoiceFolder As String
InvoiceFolder = Worksheets("AP").Range("I1").Value

Dim Singlemail, SinglemailReq As String
Singlemail = ""

'EMAIL TEXT
MessageText = "<font face=Arial>Hi Team, <br> <br> Please process attached invoice. <u><b>Please inform us when the invoice gets <font color=#3399ff>booked</font> on Your side.</b></u><br> <br></font>"

'EXPORT TO TEMP SHEET
Sheets("AP").Range("A2:K500").SpecialCells(xlCellTypeConstants).Copy Destination:=Sheets("temp").Range("A1")
Sheets("temp").Columns.AutoFit

'SET BU VARIABLE
BU = Sheets("temp").Range("I2").Value

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

'SIGNATURE
SigString = Environ("appdata") & "\Microsoft\Signatures\AP.htm"

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
    SingleEmail = 1
    SingleEmailReq = " (" + Sheets("temp").Range("A2").Value + ")"
    
Else

    Set rng = Sheets("temp").Range("A1:K500").SpecialCells(xlCellTypeConstants)
    SingleEmail = 0
    
End If

'EMAIL
With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

On Error Resume Next

With OutMail
    .To = BUmail
    .CC = "EUMarketingP2P@Staples-Solutions.com"
    .BCC = ""
    .Subject = "AP Marketing Invoice " + BU + SingleEmailReq
    .HTMLBody = MessageText + RangetoHTML(rng) + "<br> <br>" + Signature
    .SentOnBehalfOfName = "EUMarketingP2P@Staples-Solutions.com"
    .Display
End With

'ATTACHMENTS

'INV:

'-----------------CONDITION FOR UK START--------------------

If Sheets("temp").Range("I2").Value = "UK - Staples UK Online Limited" Or Sheets("temp").Range("I2").Value = "UK - Staples UK Adv Limited" Then

                    InvoiceNumber = Sheets("temp").Cells(2, 3).Value
                    InvoiceAttachment = "G:\PTP Marketing\01. Operations\08. UK Invoices\" + InvoiceNumber + ".pdf"
                    OutMail.Attachments.Add (InvoiceAttachment)
                    
                    Call ClearTempSheet
                    Exit Sub

Else

'-----------------CONDITION FOR UK END--------------------
    
        If SingleEmail = 1 Then
        
                    InvoiceNumber = Sheets("temp").Cells(2, 3).Value
                    InvoiceAttachment = "G:\PTP Marketing\01. Operations\03. Europe Marketing Invoices\" + InvoiceNumber + ".pdf"
                    OutMail.Attachments.Add (InvoiceAttachment)
        
        Else
        
            For n = 2 To 30
                    InvoiceNumber = Sheets("temp").Cells(n, 3).Value
                    InvoiceAttachment = "G:\PTP Marketing\01. Operations\03. Europe Marketing Invoices\" + InvoiceNumber + ".pdf"
                    OutMail.Attachments.Add (InvoiceAttachment)
            Next n


End If

'PO:
If SingleEmail = 1 Then

    PONumber = "PO " & Sheets("temp").Cells(2, 2).Value
    
    PONumberDir = Dir("G:\PTP Marketing\01. Operations\05. Finalised PO Folder FY 2017\" & PONumber & "*.pdf")
    POAttachment = "G:\PTP Marketing\01. Operations\05. Finalised PO Folder FY 2017\" + PONumberDir
    OutMail.Attachments.Add (POAttachment)
    
'    PONumberDir2018 = Dir("G:\PTP Marketing\01. Operations\06. Finalised PO Folder FY 2018\" & PONumber & "*.pdf")
'    POAttachment2018 = "G:\PTP Marketing\01. Operations\06. Finalised PO Folder FY 2018\" + PONumberDir2018
'    OutMail.Attachments.Add (POAttachment2018)
        
Else

    For m = 2 To 30
        
        PO = "PO "
        PONumber = Sheets("temp").Cells(m, 2).Value
        PoPdf = PO + PONumber
                
        If PONumber = "" Then Exit For
                
        PONumberDir = Dir("G:\PTP Marketing\01. Operations\05. Finalised PO Folder FY 2017\" & PoPdf & "*.pdf")
        POAttachment = "G:\PTP Marketing\01. Operations\05. Finalised PO Folder FY 2017\" + PONumberDir
        OutMail.Attachments.Add (POAttachment)
        
        
'        PONumberDir2018 = Dir("G:\PTP Marketing\01. Operations\06. Finalised PO Folder FY 2018\" & PoPdf & "*.pdf")
'        POAttachment2018 = "G:\PTP Marketing\01. Operations\06. Finalised PO Folder FY 2018\" + PONumberDir2018
'        OutMail.Attachments.Add (POAttachment2018)
                 
                 
    Next m

End If


On Error GoTo 0


On Error GoTo REMINDER

'MSG:
If SingleEmail = 1 Then

    Msg = "Invoice #"
    ReqNumber = Sheets("temp").Cells(2, 1).Value
    MsgFile = Msg + ReqNumber
    MsgAttachmentDir = Dir("G:\PTP Marketing\01. Operations\07. Approvals\" & MsgFile & "*.msg")
    MsgAttachment = "G:\PTP Marketing\01. Operations\07. Approvals\" + MsgAttachmentDir
    OutMail.Attachments.Add (MsgAttachment)

Else

    For m = 2 To 30
    
        Msg = "Invoice #"
        ReqNumber = Sheets("temp").Cells(m, 1).Value
        MsgFile = Msg + ReqNumber
        
    If ReqNumber = "" Then Exit For
    
        MsgAttachmentDir = Dir("G:\PTP Marketing\01. Operations\07. Approvals\" & MsgFile & "*.msg")
        MsgAttachment = "G:\PTP Marketing\01. Operations\07. Approvals\" + MsgAttachmentDir
        OutMail.Attachments.Add (MsgAttachment)
        
                  
    Next m

End If
End If

'CLEAR TEMP SHEET
Call ClearTempSheet

Exit Sub

'ERROR HANDLER
REMINDER:
MsgBox ("UPDATE APPROVAL FOLDER FIRST!")
Call ClearTempSheet

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



