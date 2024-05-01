# PowerShell File Integrity Checker

This PowerShell script defines a function `Get-FileProperties` to retrieve various properties and checksums of a specified file to ensure its integrity.

## Function Details

### Get-FileProperties
- This function takes a file path as input and retrieves the following properties:
  - MD5 hash
  - SHA-1 hash
  - SHA-256 hash
  - File size in bytes
  - File type (MIME type)

### Usage Example
To use this script, provide the file path of the target file to the `Get-FileProperties` function. It will calculate the checksums and retrieve the file properties.

```powershell
# Define the function
function Get-FileProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    # Check if the file exists
    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        Write-Host "File '$FilePath' not found." -ForegroundColor Red
        return
    }

    # Get file details using FileInfo object
    $fileInfo = Get-Item -Path $FilePath

    # Calculate MD5 checksum
    $md5Hash = Get-FileHash -Path $FilePath -Algorithm MD5

    # Calculate SHA-1 checksum
    $sha1Hash = Get-FileHash -Path $FilePath -Algorithm SHA1

    # Calculate SHA-256 checksum
    $sha256Hash = Get-FileHash -Path $FilePath -Algorithm SHA256

    # Get MIME type using System.IO.Path
    $mimeType = [System.IO.Path]::GetMimeType($FilePath)

    # Construct file properties object
    $properties = @{
        "File Path" = $FilePath
        "MD5" = $md5Hash.Hash.ToLower()
        "SHA-1" = $sha1Hash.Hash.ToLower()
        "SHA-256" = $sha256Hash.Hash.ToLower()
        "File Size (bytes)" = $fileInfo.Length
        "File Type" = $mimeType
    }

    return $properties
}

# Usage example
$filePath = "C:\Path\To\Your\File.ext"
$fileProperties = Get-FileProperties -FilePath $filePath

if ($fileProperties) {
    Write-Host "File Properties:"
    $fileProperties.GetEnumerator() | ForEach-Object {
        Write-Host "$($_.Key): $($_.Value)"
    }
}
