Attribute VB_Name = "OpenInvoice"
Sub OpenInvoiceFile()

   Dim PDF As Object
   Set PDF = CreateObject("Shell.Application")
   
   Dim InvNum As String
   InvNum = ActiveCell.Value
   
   
   InvoiceFile = "G:\PTP Marketing\01. Operations\03. Europe Marketing Invoices\" + InvNum + ".pdf"
   PDF.Open (InvoiceFile)
   
End Sub
