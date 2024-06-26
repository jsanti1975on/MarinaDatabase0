Option Explicit

    ' Gets the current User Name
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

Private Sub CommandButton2_Click()
    ' Search for TenantID# and display image
    Dim searchID As String
    Dim foundRow As Range
    Dim imgPath As String
    
    ' Validate TenantID# input
    searchID = Trim(TextBox3.Value)
    If searchID = "" Then
        MsgBox "Please enter TenantID# to search.", vbExclamation
        Exit Sub
    End If
    
    ' Search for TenantID# in Sheet1 column C
    Set foundRow = ThisWorkbook.Sheets("Sheet1").Columns("C").Find(What:=searchID, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Construct image path based on TenantID#
    If Not foundRow Is Nothing Then
        imgPath = "Z:\SharesHome\" & searchID & ".jpg"
        
        ' Check if image file exists
        If Dir(imgPath) <> "" Then
            ' Load and display image in Image1
            Image1.Picture = LoadPicture(imgPath)
        Else
            MsgBox "Image not found for this TenantID#.", vbExclamation
            
            ' Clear all text boxes
            TextBox1.Value = ""
            TextBox2.Value = ""
            TextBox3.Value = ""
            txtFLNumber.Value = ""
            txtPhone0.Value = ""
            txtPhone1.Value = ""
            txtEmail0.Value = ""
        End If
    Else
        MsgBox "TenantID# not found.", vbExclamation
    End If
End Sub

Private Sub CommandButton3_Click()
    ' Browse and load image into Image1
    Dim imgPath As Variant
    
    ' Set the initial directory path
    Dim initialPath As String
    initialPath = "Z:\SharesHome\"
    
    ' Open file dialog to select an image file
    imgPath = Application.GetOpenFilename("Images (*.jpg; *.jpeg; *.png),*.jpg; *.jpeg; *.png", , "Select an Image", initialPath)
    
    ' Load and display selected image in Image1 if a file is selected
    If imgPath <> False Then
        Image1.Picture = LoadPicture(imgPath)
    End If
End Sub



Private Sub TextBox3_AfterUpdate()
    ' Retrieve data based on entered TenantID#
    Dim searchID As String
    Dim foundRow As Range
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Validate TenantID# input
    searchID = Trim(TextBox3.Value)
    If searchID = "" Then
        Exit Sub
    End If
    
    ' Set worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Search for TenantID# in Sheet1 column C
    Set foundRow = ws.Columns("C").Find(What:=searchID, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If TenantID# is found, populate other text boxes with corresponding data
    If Not foundRow Is Nothing Then
        ' Get row number of the found record
        lastRow = foundRow.Row
        
        ' Populate text boxes with data from corresponding columns
        TextBox1.Value = ws.Cells(lastRow, "A").Value  ' Name
        TextBox2.Value = ws.Cells(lastRow, "B").Value  ' SlipSize
        txtFLNumber.Value = ws.Cells(lastRow, "E").Value  ' FLNumber
        txtPhone0.Value = ws.Cells(lastRow, "F").Value  ' Phone0
        txtPhone1.Value = ws.Cells(lastRow, "G").Value  ' Phone1
        txtEmail0.Value = ws.Cells(lastRow, "H").Value  ' Email0
        
        ' Optional: Display image associated with the TenantID# (if available)
        Dim imgPath As String
        imgPath = "Z:\SharesHome\" & searchID & ".jpg"
        
        If Dir(imgPath) <> "" Then
            Image1.Picture = LoadPicture(imgPath)
        Else
            Image1.Picture = LoadPicture("")  ' Clear image if not found
        End If
    Else
        ' Clear text boxes if TenantID# is not found
        TextBox1.Value = ""
        TextBox2.Value = ""
        txtFLNumber.Value = ""
        txtPhone0.Value = ""
        txtPhone1.Value = ""
        txtEmail0.Value = ""
        Image1.Picture = LoadPicture("")  ' Clear image
    End If
    
End Sub

