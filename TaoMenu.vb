Sub TaoMenu()
    Dim ws As Worksheet
    Dim x As Integer
    x = 5
    Sheets("tonghop").Range("A:A").Clear
    Sheets("tonghop").Range("A4") = "M" & ChrW(7909) & "c l" & ChrW(7909) & _
    "c các sheet trong file excel"
    For Each ws In Worksheets
        If ws.Name <> "tonghop" Then
            ws.Select
            Range("A1").Select
            ActiveSheet.Hyperlinks.Add _
            Anchor:=Selection, Address:="", SubAddress:= _
            Sheets("tonghop").Name & "!A1", TextToDisplay:="Back"
            Sheets("tonghop").Select
            Sheets("tonghop").Cells(x, 1).Select
            ActiveSheet.Hyperlinks.Add _
            Anchor:=Selection, Address:="", SubAddress:= _
            ws.Name & "!A1", TextToDisplay:=ws.Name
            x = x + 1
        End If
    Next ws
End Sub
