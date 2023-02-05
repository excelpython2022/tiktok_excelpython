Sub Compare2Columns()
Dim rngA As Range, rngB As Range, rngCell As Range
Set rngA = Range("A2:A23")
Set rngB = Range("B2:B23")
For Each rngCell In rngA
    If IsError(Application.Match(rngCell.Value, rngB, 0)) Then
        rngCell.Interior.Color = vbGreen
    End If
Next rngCell
For Each rngCell In rngB
    If IsError(Application.Match(rngCell.Value, rngA, 0)) Then
        rngCell.Interior.Color = vbYellow
    End If
Next rngCell
End Sub
