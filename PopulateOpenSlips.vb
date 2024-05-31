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
    Dim openSlipRowIndices As Collection
    Dim slipNumber As String
    Dim rowNumber As Variant ' Change to Variant to match the type stored in the collection

    ' Set the worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Initialize the collection to store the rows with "Open_Slip"
    Set openSlipRowIndices = New Collection

    ' Loop through the range A1:A81 to find "Open_Slip"
    Set rng = ws.Range("A1:A81")
    For Each cell In rng
        If cell.Value = "Open_Slip" Then
            ' Add the row number to the collection
            openSlipRowIndices.Add cell.Row
        End If
    Next cell

    ' Populate the ListBox
    Me.lstOpenSlips.Clear ' Clear any existing items

    ' Set the header for the ListBox
    Me.lstOpenSlips.AddItem "Open Slips"

    ' Add the slip numbers to the ListBox
    For Each rowNumber In openSlipRowIndices
        slipNumber = ws.Cells(rowNumber, 2).Value ' Column B contains slip number
        Me.lstOpenSlips.AddItem slipNumber
    Next rowNumber
End Sub
