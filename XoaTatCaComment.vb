Sub XoaTatCaComment()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    For Each cell In ws.UsedRange
        If Not cell.Comment Is Nothing Then cell.Comment.Delete
    Next
End Sub
