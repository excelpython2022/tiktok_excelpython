Sub Fill_Alphabet()
i = 65
For Each o In Selection
    o.Value = Chr(i)
    If i = 90 Then
        i = 64
    End If
    i = i + 1
Next o
End Sub
