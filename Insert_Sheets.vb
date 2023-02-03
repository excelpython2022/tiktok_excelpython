Sub insertsheets()
tensheet = InputBox("Nhap ten Sheet", "Thong So 1")
soluongsheet = InputBox("Nhap so luong Sheet", "Thong So 2")
For i = 1 To soluongsheet
    r = WorksheetFunction.RandBetween(0, 255)
    g = WorksheetFunction.RandBetween(0, 255)
    b = WorksheetFunction.RandBetween(0, 255)
    Sheets.Add After:=ActiveSheet, Count:=1
    ActiveSheet.Name = tensheet & i
    ActiveSheet.Range("A1") = "B" & ChrW(7841) & "n " & ChrW(273) & "ang " & ChrW(7903) & " Sheet: " & tensheet & " " & i
    ActiveSheet.Range("A1").Font.Bold = True
    ActiveSheet.Range("A1").Font.Color = vbRed
    ActiveSheet.Range("A1").Font.Size = 20
    ActiveSheet.Tab.Color = RGB(r, g, b)
Next i
End Sub
