Attribute VB_Name = "AP_DataImport"
Option Explicit

Public Sub GetAP()
Attribute GetAP.VB_ProcData.VB_Invoke_Func = "D\n14"
    Application.ScreenUpdating = False
    
Dim ExternalFile As String
Dim Answer As Integer

Answer = MsgBox("Do you want to extract data for AP Upload?", vbYesNo)
    If Answer = vbYes Then
        ExternalFile = Application.GetOpenFilename(FileFilter:="Wszystkie pliki (*.*),*.*", Title:="INVOICE DATA")
                                
        Workbooks.Open ExternalFile
    
'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData
 
'Utworzenie nowych arkuszy
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "AP UPLOAD"
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "GERMANY"

'Wlaczenie filtra dla faktur do AP
Application.ActiveWorkbook.Worksheets("DataTables").Range("A1").AutoFilter _
    Field:=8, Criteria1:="Pending Invoice Oracle AP Upload", VisibleDropDown:=True

'Skopiowanie wyników filtrowania do AP
ActiveWorkbook.Worksheets("DataTables").Range("A:M").Copy Destination:=Worksheets("AP UPLOAD").Range("A:M")

'usuniecie zbednych kolumn w arkuszu z danymi AP
Application.ActiveWorkbook.Worksheets("AP Upload").Columns("D:D").Delete
Application.ActiveWorkbook.Worksheets("AP Upload").Columns("G:G").Delete

'oczyszczenie Cost Center
Dim Cost As Range
Dim Counter As Integer

For Each Cost In Worksheets("AP Upload").Range("K2:K100")
    Counter = InStr(1, Cost, " ")
        If Counter > 0 Then
        Cost.Value = Left(Cost.Value, InStr(1, Cost.Value, " "))
        End If
    Cost.Replace What:="1•", Replacement:="", SearchOrder:=xlByColumns
Next Cost

''oczyszczenie Type of Spend
'Application.ActiveWorkbook.Worksheets("AP Upload").Columns("J").Replace _
' What:="1•", Replacement:=")", _
' SearchOrder:=xlByColumns
 
'dopasowanie szerokoœci kolumn
Application.ActiveWorkbook.Worksheets("AP Upload").Columns("A:I").AutoFit

'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData

'GERMANY

'Deklaracja tablicy dla filtrowania
Dim Germany As Variant
Germany = Array("DE - SE Retail Germany", "DE - SE Advantage Germany", "DE - SE SBD Germany Total")

'Filtrowanie
Application.ActiveWorkbook.Worksheets("DataTables").Range("C1").AutoFilter _
    Field:=11, _
    Criteria1:=(Germany), _
    Operator:=xlFilterValues, _
    VisibleDropDown:=True

Application.ActiveWorkbook.Worksheets("DataTables").Range("C1").AutoFilter _
    Field:=8, _
    Criteria1:="Pending Invoice Oracle AP Upload", _
    VisibleDropDown:=True
    
'Kopiowanie do nowej karty
ActiveWorkbook.Worksheets("DataTables").Range("A:R").Copy Destination:=Worksheets("GERMANY").Range("A:M")

'dopasowanie szerokoœci kolumn
Application.ActiveWorkbook.Worksheets("GERMANY").Columns("K").AutoFit


'CLEAR FILTERS
If ActiveWorkbook.Worksheets("DataTables").FilterMode = True _
    Then Worksheets("DataTables").ShowAllData
    
Worksheets("AP Upload").Activate


    Else
        'Do nothing
    End If
    

MsgBox ("Task completed!")


End Sub


