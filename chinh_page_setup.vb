Sub AdjustPageSetup()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        ws.PageSetup.Zoom = 100
        ActiveWindow.View = 2
        If ws.VPageBreaks.Count >= 1 Then
          ws.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        End If
        ActiveWindow.View = 1
   Next ws
End Sub
