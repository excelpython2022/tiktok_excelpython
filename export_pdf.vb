Sub SaveWorkshetAsPDF()
Dim ws As Worksheet
For Each ws In Worksheets
    ws.ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & "\" & ws.Name & ".pdf"
Next ws
End Sub