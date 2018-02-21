Attribute VB_Name = "RefreshFilters"
Sub ReFilterAP()

If Workbooks("APFA.xlsm").Worksheets("AP").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("AP").ShowAllData
    
Workbooks("APFA.xlsm").Worksheets("AP").Range("A1").AutoFilter Field:=16, Criteria1:="", VisibleDropDown:=True
Workbooks("APFA.xlsm").Worksheets("AP").Range("B1").AutoFilter Field:=2, Criteria1:=">0", VisibleDropDown:=True

End Sub

Sub ReFilterFA()

If Workbooks("APFA.xlsm").Worksheets("FA").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("FA").ShowAllData

Workbooks("APFA.xlsm").Worksheets("FA").Range("A1").AutoFilter Field:=15, Criteria1:="", VisibleDropDown:=True

End Sub


Sub ClearFilters()

If Workbooks("APFA.xlsm").Worksheets("AP").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("AP").ShowAllData

End Sub



