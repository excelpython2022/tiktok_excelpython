Sub ExportSelectionToPDF()
Dim FileName As String
Dim ws As Worksheet
Set ws = ActiveSheet
FileName = Application.GetSaveAsFilename(InitialFileName:=ws.Name, _
FileFilter:="PDF Files (*.pdf), *.pdf")
If FileName = "False" Then Exit Sub
With ActiveSheet.PageSetup
    .PrintArea = ActiveSheet.Cells(1, 1).CurrentRegion.Address
    .Zoom = False
    .Orientation = xlLandscape
    .FitToPagesWide = 1
    .FitToPagesTall = 1
End With
Selection.ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName, Quality:=xlQualityStandard
End Sub
