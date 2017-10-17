Attribute VB_Name = "RefreshFilterFA"
Sub ReFilterFA()

If Workbooks("APFA.xlsm").Worksheets("FA").FilterMode = True _
    Then Workbooks("APFA.xlsm").Worksheets("FA").ShowAllData

Workbooks("APFA.xlsm").Worksheets("FA").Range("A1").AutoFilter Field:=11, Criteria1:="", VisibleDropDown:=True

End Sub


