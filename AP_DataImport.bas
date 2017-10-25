Attribute VB_Name = "AP_DataImport"
Option Explicit

Public Sub GetAP()
Attribute GetAP.VB_ProcData.VB_Invoke_Func = "D\n14"

Application.ScreenUpdating = False
    
Dim ExternalFile As String
Dim Answer As Integer
Dim Cost As Range
Dim Counter As Integer
Dim Germany As Variant
Dim FA As Range
Dim CounterFa As Integer
Dim LastRowAP As String
Dim LastRowFA As String
Dim InvoiceData As Workbook

'GET USER CONFIRMATION
Answer = MsgBox("Do you want to extract data for AP/FA Upload?", vbYesNo)
    If Answer = vbYes Then
        ExternalFile = Application.GetOpenFilename(FileFilter:="Wszystkie pliki (*.*),*.*", Title:="INVOICE DATA")
                                
        Workbooks.Open ExternalFile
        
        Set InvoiceData = ActiveWorkbook
    
'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData

'NEW WORSHEETS
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "AP UPLOAD"
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "FA UPLOAD"
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "GERMANY"

'SET FILTER (AP)
Application.ActiveWorkbook.Worksheets("DataTables").Range("A1").AutoFilter _
    Field:=8, Criteria1:="Pending Invoice Oracle AP Upload", VisibleDropDown:=True

'COPY TO AP SHEET
ActiveWorkbook.Worksheets("DataTables").Range("A:M").Copy Destination:=Worksheets("AP UPLOAD").Range("A:M")

'DELETING COLUMNS
With Application.ActiveWorkbook.Worksheets("AP Upload")
    .Columns("D:D").Delete
    .Columns("G:G").Delete
    .Range("A1:K1").Delete
End With

'CLEAN UP IN COLUMN K
For Each Cost In Worksheets("AP Upload").Range("K1:K100")
    Counter = InStr(1, Cost, " ")
        If Counter > 0 Then
        Cost.Value = Left(Cost.Value, InStr(1, Cost.Value, " "))
        End If
    Cost.Replace What:="1•", Replacement:="", SearchOrder:=xlByColumns
Next Cost

''CLEAN UP IN COLUMN J
'Application.ActiveWorkbook.Worksheets("AP Upload").Columns("J").Replace _
'What:="1•", Replacement:=")", _
'SearchOrder:=xlByColumns
 
'AUTOFIT COLUMNS
Application.ActiveWorkbook.Worksheets("AP Upload").Columns("A:I").AutoFit

'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
Then Worksheets("DataTables").ShowAllData

'ARRAY & FILTER (DE)
Germany = Array("DE - SE Retail Germany", "DE - SE Advantage Germany", "DE - SE SBD Germany Total")

Application.ActiveWorkbook.Worksheets("DataTables").Range("C1").AutoFilter _
    Field:=11, _
    Criteria1:=(Germany), _
    Operator:=xlFilterValues, _
    VisibleDropDown:=True

Application.ActiveWorkbook.Worksheets("DataTables").Range("C1").AutoFilter _
    Field:=8, _
    Criteria1:="Pending Invoice Oracle AP Upload", _
    VisibleDropDown:=True
    
'COPY TO EMPTY SHEET
ActiveWorkbook.Worksheets("DataTables").Range("A:R").Copy Destination:=Worksheets("GERMANY").Range("A:M")

'AUTOFIT
Application.ActiveWorkbook.Worksheets("GERMANY").Columns("K").AutoFit

'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData

'SET FILTER (FA)
Application.ActiveWorkbook.Worksheets("DataTables").Range("B1").AutoFilter _
    Field:=8, Criteria1:="Pending Invoice Oracle FA Upload", VisibleDropDown:=True

'COPY TO EMPTY SHEET
ActiveWorkbook.Worksheets("DataTables").Range("A:M").Copy Destination:=Worksheets("FA UPLOAD").Range("A:M")

'DELETE COLUMNS
With Application.ActiveWorkbook.Worksheets("FA Upload")
    .Columns("K:K").Delete
    .Columns("H:H").Delete
    .Columns("D:D").Delete
    .Range("A1:J1").Delete
End With

'CLEAN UP IN COLUMN J
For Each FA In Worksheets("FA Upload").Range("J1:J100")
    CounterFa = InStr(1, FA, " ")
        If CounterFa > 0 Then
        FA.Value = Left(FA.Value, InStr(1, FA.Value, " "))
        End If
    FA.Replace What:="1•", Replacement:="", SearchOrder:=xlByColumns
Next FA

'AUTOFIT
Application.ActiveWorkbook.Worksheets("FA Upload").Columns("D:H").AutoFit

'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData

'IMPORT TO THIS WORKBOOK
On Error Resume Next

LastRowAP = Workbooks("APFA.xlsm").Worksheets("AP").Cells(Rows.Count, "A").End(xlUp).Row + 1
LastRowFA = Workbooks("APFA.xlsm").Worksheets("FA").Cells(Rows.Count, "A").End(xlUp).Row + 1

Worksheets("AP UPLOAD").Range("A:K").SpecialCells(xlCellTypeConstants).Copy
Workbooks("APFA.xlsm").Worksheets("AP").Range("A" & LastRowAP & ":K" & LastRowAP).PasteSpecial xlPasteValues

If Worksheets("FA UPLOAD").Range("A1") > 0 Then

Worksheets("FA UPLOAD").Range("A:J").SpecialCells(xlCellTypeConstants).Copy
Workbooks("APFA.xlsm").Worksheets("FA").Range("A" & LastRowFA & ":J" & LastRowFA).PasteSpecial xlPasteValues

Else
End If

'COPY DATA SHEET TO THIS WORKBOOK
Application.DisplayAlerts = False

Workbooks("APFA.xlsm").Worksheets("DataTables").Delete
InvoiceData.Worksheets("DataTables").Copy After:=Workbooks("APFA.xlsm").Worksheets("temp")

Application.DisplayAlerts = True

Worksheets("AP").Activate

On Error GoTo 0

    Else
    
    MsgBox ("Task aborted!")
    
    End If
    
Application.ScreenUpdating = True

End Sub


