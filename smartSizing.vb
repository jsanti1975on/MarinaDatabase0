    ' Testing
Private Sub UserForm_Initialize()
    ' Load and display an image in Image1 with smart sizing
    Dim imgPath As String
    imgPath = "C:/Path/ToYour/Image.jpg"  ' Path to your image file
    
    ' Check if the image file exists
    If Dir(imgPath) <> "" Then
        ' Set the PictureSizeMode to control how the image is displayed
        Me.Image1.PictureSizeMode = fmPictureSizeModeStretch  ' Adjust as needed
        
        ' Load and display the image in Image1
        Me.Image1.Picture = LoadPicture(imgPath)
    Else
        MsgBox "Image not found.", vbExclamation
    End If
End Sub
