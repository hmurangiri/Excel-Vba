' This is a script to combine several workbooks with the same populated columns in one sheet
' The asumption is that the data is the same columns and that the data is in the first sheet and that every sheet has data with the same header
    
Dim count
Dim fileCount

Sub CombineDataFromSelectedWorkbooks()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = True
        .Title = "Select Excel Files"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.csv"

        If .Show = True Then
            Dim i As Integer
            fileCount = .SelectedItems.count
            count = 0
            For i = 1 To .SelectedItems.count
                Dim filePath As String
                filePath = .SelectedItems(i)
                count = count + 1
                AppendDataFromWorkbook filePath
            Next i
        End If
    End With
End Sub

Sub AppendDataFromWorkbook(filePath As String)
    Dim srcWorkbook As Workbook
    Set srcWorkbook = Workbooks.Open(filePath, True, True) ' Open the source workbook in read-only mode
    Dim copyFrom

    Dim srcSheet As Worksheet
    Set srcSheet = srcWorkbook.Sheets(1) ' Assumes data is on the first sheet

    ' Find the last row with data on the source sheet
    Dim lastRowSrc As Long
    lastRowSrc = srcSheet.Cells(srcSheet.Rows.count, 1).End(xlUp).Row

    ' Find the next empty row in the destination sheet
    Dim destSheet As Worksheet
    Set destSheet = ThisWorkbook.Worksheets("Sheet2") ' Change "Sheet2" to your master sheet's name
    
    If count = 1 Then destSheet.Columns("A:B").Delete
    
    Dim nextRowDest As Long
    nextRowDest = destSheet.Cells(destSheet.Rows.count, 1).End(xlUp).Row + 1
    
    nextRowDest = IIf(nextRowDest = 2, 1, nextRowDest)
    copyFrom = IIf(nextRowDest = 1, "A1:D", "A2:D")

    ' Copy the data to the master sheet
    srcSheet.Range(copyFrom & lastRowSrc).Copy destSheet.Range("A" & nextRowDest)

    ' Close the source workbook
    srcWorkbook.Close False
    
    ' Delete unnecessary columns
    If count = fileCount Then
        MsgBox "File combination complete!"
    End If
End Sub
