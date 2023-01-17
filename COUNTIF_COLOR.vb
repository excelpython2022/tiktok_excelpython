Function COUNTIF_COLOR(rng1 As Range, rng2 As Range) As Long
Dim tong As Long
Dim o As Range
tong = 0
For Each o In rng1
    If o.Interior.Color = rng2.Interior.Color Then
        tong = tong + 1
    End If
Next
COUNTIF_COLOR = tong
End Function
