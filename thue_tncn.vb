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
