Private Sub CommandButton1_Click()
    ' Save data to worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim slipNumberValue As String
    Dim existingRow As Range
    
    ' Set worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Validate required input fields
    If Trim(TextBox1.Value) = "" Or Trim(TextBox2.Value) = "" Or Trim(TextBox3.Value) = "" Then
        MsgBox "Please enter Name, Slip Number, and TenantID#.", vbExclamation
        Exit Sub
    End If
    
    ' Get SlipNumber value from TextBox2
    slipNumberValue = Trim(TextBox2.Value)
    
    ' Check if SlipNumber exceeds 80
    If Val(slipNumberValue) > 80 Then
        MsgBox "Error Code 1: This application is based on 80 slips. Please enter a Slip Number between 1 and 80.", vbExclamation
        Exit Sub
    End If
    
    ' Check for duplicate SlipNumber in Column B
    Set existingRow = ws.Columns("B").Find(What:=slipNumberValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If duplicate SlipNumber is found, display message and exit sub
    If Not existingRow Is Nothing Then
        MsgBox "Duplicate Slip Number found. Please enter a different Slip Number.", vbExclamation
        Exit Sub
    End If
    
    ' Find the next available row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Write data to worksheet
    With ws
        .Cells(lastRow, "A").Value = Trim(TextBox1.Value)     ' Name
        .Cells(lastRow, "B").Value = slipNumberValue          ' SlipNumber
        .Cells(lastRow, "C").Value = Trim(TextBox3.Value)     ' TenantID#
        .Cells(lastRow, "E").Value = Trim(txtFLNumber.Value)  ' FLNumber
        .Cells(lastRow, "F").Value = Trim(txtPhone0.Value)    ' Phone0
        .Cells(lastRow, "G").Value = Trim(txtPhone1.Value)    ' Phone1
        .Cells(lastRow, "H").Value = Trim(txtEmail0.Value)    ' Email0
        .Cells(lastRow, "I").Value = Application.UserName      ' Log current user
        .Cells(lastRow, "J").Value = Format(Now(), "DD-MM-YYYY HH:MM:SS") ' Log timestamp
    End With
    
    ' Clear input controls after saving
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    txtFLNumber.Value = ""
    txtPhone0.Value = ""
    txtPhone1.Value = ""
    txtEmail0.Value = ""
    
    ' Clear image in Image1
    Image1.Picture = LoadPicture("")
    
    ' Display success message
    MsgBox "Data saved successfully.", vbInformation
End Sub
