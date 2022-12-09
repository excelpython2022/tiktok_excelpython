Sub start_nhapnhay()
    mau_chu_1 = WorksheetFunction.RandBetween(1, 56)
    mau_nen_1 = WorksheetFunction.RandBetween(1, 56)
    Do While mau_chu_1 = mau_nen_1
        mau_nen_1 = WorksheetFunction.RandBetween(1, 56)
    Loop
    mau_chu_2 = WorksheetFunction.RandBetween(1, 56)
    mau_nen_2 = WorksheetFunction.RandBetween(1, 56)
    Do While mau_chu_2 = mau_nen_2
        mau_nen_2 = WorksheetFunction.RandBetween(1, 56)
    Loop
    Range("C2").Font.ColorIndex = mau_chu_1
    Range("C2").Interior.ColorIndex = mau_nen_1
    Range("C3").Font.ColorIndex = mau_chu_2
    Range("C3").Interior.ColorIndex = mau_nen_2
    ThoiGian = Now + TimeSerial(0, 0, 1)
    Application.OnTime ThoiGian, "'" & ThisWorkbook.Name & "'!start_nhapnhay", , True
End Sub
Sub stop_nhapnhay()
    On Error Resume Next
    For i = 0 To 5
        Application.OnTime Now + TimeValue("00:00:" & i), "start_nhapnhay", , False
    Next i
    On Error GoTo 0
End Sub
