Sub SaveWorkshetAs_Workbook()
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate
    Set wb = Workbooks.Add
    ThisWorkbook.Activate
    ActiveSheet.Copy Before:=wb.Sheets(1)
    wb.Activate
    wb.SaveAs ThisWorkbook.Path & "\" & ws.Name & ".xlsx"
    wb.Close
Next ws
End Sub
