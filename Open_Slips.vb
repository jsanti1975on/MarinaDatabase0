Private Sub cmdLoadOpenSlips_Click()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    ' Set the worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Clear the ListBox before loading data
    lstOpenSlips.Clear

    ' Find the last row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Set the range to search through
    Set rng = ws.Range("A1:A" & lastRow)

    ' Loop through each cell in the range
    For Each cell In rng
        ' Check if the cell contains "Open_Slip"
        If cell.Value = "Open_Slip" Then
            ' Add the corresponding Slip Number from Column B to the ListBox
            lstOpenSlips.AddItem ws.Cells(cell.Row, "B").Value
        End If
    Next cell
End Sub

Private Sub UserForm_Initialize()
    ' Optionally, you can load the data automatically when the form initializes
    cmdLoadOpenSlips_Click
End Sub
