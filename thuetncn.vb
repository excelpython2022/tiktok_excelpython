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
