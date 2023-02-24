Function RemoveSpaces(ByVal s As String) As String
    Dim result As String
    s = Trim(s)
    result = ""
    Dim i As Integer
    For i = 1 To Len(s)
        If Mid(s, i, 1) = " " Then
            If i > 1 And i < Len(s) And Mid(s, i - 1, 1) <> " " Then
                result = result & " "
            End If
        Else
            result = result & Mid(s, i, 1)
        End If
    Next i
    RemoveSpaces = result
End Function
