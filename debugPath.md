```vb
' Construct image path based on TenantID#
imgPath = "C:/Users/santj/ID/" & searchID & ".jpg"
Debug.Print "Image Path: " & imgPath

' Check if image file exists
If Dir(imgPath) <> "" Then
    ' Load and display image in Image1
    Image1.Picture = LoadPicture(imgPath)
Else
    MsgBox "Image not found for this TenantID#.", vbExclamation
End If
This code block constructs an image path based on a `searchID` value and checks if the image file exists at the specified path. If the image exists, it loads and displays it in `Image1`. If the image file is not found, it displays a message box indicating that the image is not available for the specified `TenantID#`.
