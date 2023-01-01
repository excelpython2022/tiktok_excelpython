Sub SoSanh2Sheet()
    Dim rsheet1 As Range
    Dim rsheet2 As Range
    Set rsheet1 = Worksheets("Sheet1").Range("A3:AF36")
    Set rsheet2 = Worksheets("Sheet2").Range("A3:AF36")
    tongthaydoi = 0
    For Each cel In rsheet1
        Debug.Print "Searching... " & cel.Address(0, 0)
        If cel.Value <> Worksheets("Sheet2").Range(cel.Address) Then
            Worksheets("Sheet2").Range(cel.Address).Interior.Color = vbYellow
            tongthaydoi = tongthaydoi + 1
        End If
    Next cel
    Application.Assistant.DoAlert "ThÙng b·o", "CÛ t" & ChrW(7845) & "t c" & ChrW(7843) & _
        " " & tongthaydoi & " Ù thay " & ChrW(273) & ChrW(7893) & _
        "i trÍn Sheet 2", 0, 4, 0, 0, 0
End Sub
