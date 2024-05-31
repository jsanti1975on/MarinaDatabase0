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
    Dim rowNumber As Variant
    
    ' Set the worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Initialize the collection to store the rows with "Open_Slip"
    Set openSlipRowIndices = New Collection
    
    ' Loop through the range A1:A81 to find "Open_Slip"
    Set rng = ws.Range("A1:A81")
    For Each cell In rng
        If cell.Value = "Open_Slip" Then ' Confirm the correct string
            ' Add the row number to the collection
            openSlipRowIndices.Add cell.Row
        End If
    Next cell
    
    ' Loop through the collection and populate the ListBox
    For Each rowNumber In openSlipRowIndices
        slipNumber = ws.Cells(rowNumber, 2).Value ' Column B contains Slip Number
        Me.lstOpenSlips.AddItem slipNumber
    Next rowNumber
End Sub
