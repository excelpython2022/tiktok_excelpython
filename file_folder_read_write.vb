Sub TaoThuMuc()
    Dim tenThuMuc As String
    tenThuMuc = "C:\VBA"
    MkDir tenThuMuc
End Sub

Sub XoaThuMuc()
    Dim tenThuMuc As String
    tenThuMuc = "C:\VBA2"
    'mystr = "_""-""_"
    '2 dau "" thi con 1 "
    If Dir(tenThuMuc, vbDirectory) <> "" Then
        Shell "cmd /c rd """ & tenThuMuc & """", vbHide
    Else
        MsgBox "Not found " & tenThuMuc
    End If
End Sub
'Tham chieu den Microsoft Scripting Runtime
'csv va txt
Sub ReadUnicodeCSVFile()
    Dim fso As FileSystemObject
    Dim fileStream As TextStream
    Dim filePath As String
    Dim fileContent As String
    Dim rows As Variant
    Dim i As Long, j As Long
    
    Set fso = New FileSystemObject
    filePath = "C:\VBA\data1.csv"
    Set fileStream = fso.OpenTextFile(filePath, ForReading, True, TristateUseDefault)
    fileContent = fileStream.ReadAll
    rows = Split(fileContent, vbCrLf)
    
    For i = 0 To UBound(rows)
        Dim cols As Variant
        cols = Split(rows(i), ";")
        For j = 0 To UBound(cols)
            Sheet1.Cells(i + 1, j + 1).Value = cols(j)
        Next j
    Next i
    
    fileStream.Close
    Set fileStream = Nothing
    Set fso = Nothing
End Sub
Sub SaveSheetAsUnicodeCSV()
    Dim fso As FileSystemObject
    Dim fileStream As TextStream
    Dim filePath As String
    Dim fileContent As String
    Dim i As Long, j As Long
    
    Set fso = New FileSystemObject
    filePath = "C:\VBA\data1.csv"
    Set fileStream = fso.CreateTextFile(filePath, True, True)
    
    For i = 1 To 3
        Dim rowContent As String
        For j = 1 To 4
            If j > 1 And j < 4 Then rowContent = rowContent & ","
            rowContent = rowContent & Sheet1.Cells(i, j).Value
        Next j
        fileContent = fileContent & rowContent & vbCrLf
        rowContent = ""
    Next i
    
    fileStream.Write (fileContent)
    fileStream.Close
    Set fileStream = Nothing
    Set fso = Nothing
End Sub

