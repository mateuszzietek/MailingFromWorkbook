Attribute VB_Name = "CreateMail"
Sub PrepareMail()

Dim rng As Range
Dim OutApp As Object
Dim OutMail As Object

'Set rng = Nothing
'Only send the visible cells in the selection.
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
Set rng = ActiveSheet.Range("A1:K10").SpecialCells(xlCellTypeVisible)

'If rng Is Nothing Then
'    MsgBox "The selection is not a range or the sheet is protected. " & _
           vbNewLine & "Please correct and try again.", vbOKOnly
'    Exit Sub
'End If

'With Application
'    .EnableEvents = False
'    .ScreenUpdating = False
'End With

With OutMail
    .To = "mateusz.zietek@staples-solutions.com"
    .CC = "EUMarketing@Staples-solutions.com; Przemyslaw.Luczak@staples-solutions.com"
    .BCC = ""
    .Subject = "AP Marketing Invoice"
    .Body = "blabla"
End With


'With Application
'    .EnableEvents = True
'    .ScreenUpdating = True
'End With

End Sub

