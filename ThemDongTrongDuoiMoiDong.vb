Sub ThemDongTrongDuoiMoiDong()
CountRow = Selection.EntireRow.Count
For i = 1 To CountRow
    ActiveCell.EntireRow.Insert
    ActiveCell.Offset(2, 0).Select
Next i
End Sub
