Attribute VB_Name = "RefreshFilterAP"
Sub ReFilterAP()

If Workbooks("APFA.xlsm").Worksheets("AP").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("AP").ShowAllData
    
Workbooks("APFA.xlsm").Worksheets("AP").Range("A1").AutoFilter Field:=12, Criteria1:="", VisibleDropDown:=True

End Sub
