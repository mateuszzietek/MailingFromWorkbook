Attribute VB_Name = "ClearTemp"
Sub ClearTempSheet()

Sheets("temp").Range("A:Q").Delete


With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set OutMail = Nothing
Set OutApp = Nothing


End Sub
