Private Sub UserForm_Initialize()
    ' Initialize the ListBox properties
    With Me.lstOpenSlips
        .ColumnCount = 1
        .ColumnHeads = True
        .ColumnWidths = "100"
    End With

    ' Call the function to populate the ListBox
    PopulateOpenSlips
End Sub

Private Sub PopulateOpenSlips()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim openSlipRows As Collection
    Dim slipNumber As String

    ' Set the worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Initialize the collection to store rows with "Open_Slip"
    Set openSlipRows = New Collection

    ' Loop through the range A1:A81 to find "Open_Slip"
    Set rng = ws.Range("A1:A81")
    For Each cell In rng
        If cell.Value = "Open_Slip" Then
            ' Add the row number to the collection
            openSlipRows.Add cell.Row
        End If
    Next cell

    ' Loop through the collection and populate the ListBox
    For Each cell In openSlipRows
        slipNumber = ws.Cells(cell, 2).Value ' Column B contains Slip Number
        Me.lstOpenSlips.AddItem slipNumber
    Next cell
End Sub
