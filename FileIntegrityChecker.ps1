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

    # Determine file type (MIME type)
    $mimeType = ""
    $fileExtension = [System.IO.Path]::GetExtension($FilePath)
    switch -Wildcard ($fileExtension) {
        "*.docx" { $mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        "*.xlsx" { $mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        default { $mimeType = "unknown" }
    }

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

# Usage example - Replace $filePath with your desired file path
$filePath = "D:\Sec+\Microsoft Word Documents\CompTIA Security+.docx"
$fileProperties = Get-FileProperties -FilePath $filePath

if ($fileProperties) {
    # Define output file path
    $outputFilePath = Join-Path (Split-Path -Path $filePath) "testinghash.txt"

    # Export file properties to text file
    $fileProperties.GetEnumerator() | ForEach-Object {
        "$($_.Key): $($_.Value)"
    } | Out-File -FilePath $outputFilePath -Append

    Write-Host "File properties exported to: $outputFilePath"
}
