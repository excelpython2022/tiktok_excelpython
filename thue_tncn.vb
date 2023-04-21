Function THUE(tntt)
If tntt > 80000000 Then
    THUE = tntt * 0.35 - 9850000
ElseIf tntt > 52000000 Then
    THUE = tntt * 0.3 - 5850000
ElseIf tntt > 32000000 Then
    THUE = tntt * 0.25 - 3250000
ElseIf tntt > 18000000 Then
    THUE = tntt * 0.2 - 1650000
ElseIf tntt > 10000000 Then
    THUE = tntt * 0.15 - 750000
ElseIf tntt > 5000000 Then
    THUE = tntt * 0.1 - 250000
ElseIf tntt > 0 Then
    THUE = tntt * 0.05
Else
    THUE = 0
End If
End Function

Sub ThueTNCN()
x = 100
bacthue = Array(0, 5, 10, 18, 32, 52, 80, 1000000000)
phantramthue = Array(0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35)
vitri = Application.Match(x, bacthue, True)
thue = 0
For i = 1 To vitri
    If bacthue(i) <= x Then
        thue = thue + (bacthue(i) - bacthue(i - 1)) * phantramthue(i)
    Else
        thue = thue + (x - bacthue(i - 1)) * phantramthue(i)
    End If
Next i
Debug.Print thue
End Sub

Function fn_ThueTNCN(x)
'x = 100
bacthue = Array(0, 5, 10, 18, 32, 52, 80, 1000000000)
phantramthue = Array(0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35)
vitri = Application.Match(x, bacthue, True)
thue = 0
For i = 1 To vitri
    If bacthue(i) <= x Then
        thue = thue + (bacthue(i) - bacthue(i - 1)) * phantramthue(i)
    Else
        thue = thue + (x - bacthue(i - 1)) * phantramthue(i)
    End If
Next i
Debug.Print thue
fn_ThueTNCN = thue
End Function
