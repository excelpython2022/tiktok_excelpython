Function GopDong(cellValue As String) As String
    Dim arr() As String
    Dim i As Long
    arr = Split(cellValue, vbLf)
    For i = 0 To UBound(arr)
        arr(i) = Trim(arr(i))
        If i = 0 Then
            GopDong = arr(i)
        Else
            GopDong = GopDong & " " & arr(i)
        End If
    Next i
End Function
