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
