Option Explicit

 ' Function inside the UserForm1 module for error handling TextBox 3, referance error handling for TextBox 1 and 2
Private Function IsValidTenantID(slipNumber As String, tenantID As String) As Boolean
    ' Check if TenantID matches the specified format
    IsValidTenantID = False

    ' Validate length
    If Len(tenantID) <> 4 Then Exit Function

    ' Validate first two digits
    Dim firstTwoDigits As String
    firstTwoDigits = Left(tenantID, 2)
    If Not IsNumeric(firstTwoDigits) Then Exit Function
    If Not (firstTwoDigits >= "01" And firstTwoDigits <= "80") Then Exit Function

    ' Validate that the first two digits of TenantID match the SlipNumber
    If firstTwoDigits <> slipNumber Then Exit Function

    ' Validate last two digits
    Dim lastTwoDigits As String
    lastTwoDigits = Right(tenantID, 2)
    If Not IsNumeric(lastTwoDigits) Then Exit Function
    If lastTwoDigits <> "01" Then Exit Function

    IsValidTenantID = True
End Function

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim slipNumberValue As String
    Dim tenantIDValue As String
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
    tenantIDValue = Trim(TextBox3.Value)

    ' Validate SlipNumber
    If Val(slipNumberValue) > 80 Then
        MsgBox "Error Code 1: This application is based on 80 slips. Please enter a Slip Number between 1 and 80.", vbExclamation
        Exit Sub
    End If

    ' Validate TenantID format and match with SlipNumber
    If Not IsValidTenantID(slipNumberValue, tenantIDValue) Then
        MsgBox "Invalid Format: The first two digits in both the Slip# and TenantID# should be the same value (e.g. 2101 and 21)", vbExclamation
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
        .Cells(lastRow, "I").Value = Application.UserName     ' Log current user
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

Private Sub cmdUpdateTenantData_Click()
    Dim ws As Worksheet
    Dim searchID As String
    Dim foundRow As Range
    Dim response As VbMsgBoxResult

    ' Set worksheet object
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Validate TenantID# input
    searchID = Trim(txtUpdatedTenantData.Value)
    If searchID = "" Then
        MsgBox "Please enter TenantID# to update.", vbExclamation
        Exit Sub
    End If

    ' Search for TenantID# in Sheet1 column C
    Set foundRow = ws.Columns("C").Find(What:=searchID, LookIn:=xlValues, LookAt:=xlWhole)

    ' If TenantID# is found, prompt for update
    If Not foundRow Is Nothing Then
        response = MsgBox("TenantID# found. Would you like to update the tenant information?", vbYesNo + vbQuestion, "Update Tenant Information")

        If response = vbYes Then
            ' Clear input controls to allow new tenant information
            TextBox1.Value = ""
            TextBox2.Value = ""
            TextBox3.Value = ""
            txtFLNumber.Value = ""
            txtPhone0.Value = ""
            txtPhone1.Value = ""
            txtEmail0.Value = ""

            ' Set TenantID and SlipNumber in TextBox3 and TextBox2 respectively
            TextBox3.Value = searchID
            TextBox2.Value = ws.Cells(foundRow.Row, "B").Value
        End If
    Else
        MsgBox "TenantID# not found.", vbExclamation
    End If
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
        imgPath = "F:\PropShop Accounting\CUSTOMER SERVICE SPECIALIST\IDs\" & searchID & ".jpg"

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
    initialPath = "F:\PropShop Accounting\CUSTOMER SERVICE SPECIALIST\IDs"

    ' Open file dialog to select an image file
    imgPath = Application.GetOpenFilename("Images (*.jpg; *.jpeg; ),*.jpg; *.jpeg; ", , "Select an Image", initialPath)

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

    ' Check if searchID is valid
    If Not IsValidTenantID(Left(searchID, 2), searchID) Then
        Call ShowTenantIDFormatMessage
        Image1.Picture = LoadPicture("")  ' Clear image if not found
        On Error GoTo 0
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

Private Sub ShowTenantIDFormatMessage()
    ' Define the message to display when the TenantID format is invalid
    MsgBox "Invalid TenantID format. TenantID should be in the format Slip#01 and match the Slip Number.", vbExclamation, "Invalid TenantID"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Opens the application with Microsoft Excel running in background
    Application.Visible = True
    
    Unload Me
End Sub
