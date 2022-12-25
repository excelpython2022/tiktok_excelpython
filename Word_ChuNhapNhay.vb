Sub Word_ChuNhapNhay()
    Dim arr As Variant
    arr = Array(0, 1, 2, 4, 9, 13, 14, 16, 15, 11, 0, 5, 6, 10, 3, 12, 8, 7)
    mauchu = Int((17 * Rnd) + 1)
    Debug.Print mauchu
    Selection.Font.ColorIndex = arr(mauchu)
    Application.OnTime When:=Now + TimeValue("00:00:01"), Name:="Word_ChuNhapNhay"
End Sub
