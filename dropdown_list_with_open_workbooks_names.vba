' This is a script to populate a given cell with a list of workbook names for the open workbooks
Sub PopulateDropdownWithOpenWorkbooks()
    Dim dropdownCell As Range
    Set dropdownCell = ThisWorkbook.Sheets("Sheet1").Range("C2") ' Update with your sheet name and cell address
    
    dropdownCell.ClearContents

    Dim openWorkbook As Workbook
    Dim workbookNames() As String
    Dim i As Integer
    i = 0

    For Each openWorkbook In Application.Workbooks
        ReDim Preserve workbookNames(i)
        workbookNames(i) = openWorkbook.Name
        i = i + 1
    Next openWorkbook

    With dropdownCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Join(workbookNames, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub
