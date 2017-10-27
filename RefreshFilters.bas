Attribute VB_Name = "RefreshFilters"
Sub ReFilterAP()

If Workbooks("APFA.xlsm").Worksheets("AP").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("AP").ShowAllData
    
Workbooks("APFA.xlsm").Worksheets("AP").Range("A1").AutoFilter Field:=12, Criteria1:="", VisibleDropDown:=True

End Sub

Sub ReFilterFA()

If Workbooks("APFA.xlsm").Worksheets("FA").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("FA").ShowAllData

Workbooks("APFA.xlsm").Worksheets("FA").Range("A1").AutoFilter Field:=11, Criteria1:="", VisibleDropDown:=True

End Sub


