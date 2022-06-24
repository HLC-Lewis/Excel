'Copy a sheet based on a list

Sub CopySheet()

Dim i As Long, LastRow As Long, ws As Worksheet
Sheets("Learners").Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To LastRow
    Sheets("Template").Copy After:=Sheets(i)
    ActiveSheet.Name = Sheets("Learners").Cells(i, 1)
Next i

End Sub
