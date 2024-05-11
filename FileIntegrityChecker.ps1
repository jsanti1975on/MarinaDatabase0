# Define the function to get file properties
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

    # Calculate MD5 checksum
    $md5Hash = Get-FileHash -Path $FilePath -Algorithm MD5

    # Calculate SHA-256 checksum
    $sha256Hash = Get-FileHash -Path $FilePath -Algorithm SHA256

    # Get only the file name from the full path
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

    # Construct file properties object
    $properties = @{
        "File Name" = $fileName
        "MD5" = $md5Hash.Hash.ToLower()
        "SHA-256" = $sha256Hash.Hash.ToLower()
    }

    return $properties
}

# Specify the directory path to scan for .jpg files (replace with your desired directory)
$directoryPath = "D:\IDs"

# Get all .jpg files in the specified directory
$files = Get-ChildItem -Path $directoryPath -File -Filter *.jpg

# Array to store file properties
$filePropertiesList = @()

# Loop through each .jpg file and get its properties
foreach ($file in $files) {
    $fileProperties = Get-FileProperties -FilePath $file.FullName
    if ($fileProperties) {
        $filePropertiesList += $fileProperties
    }
}

# Define the output file path (same directory as the IDs directory)
$outputFilePath = Join-Path (Split-Path -Path $directoryPath) "hashes.txt"

# Output file properties to the formatted text file (hashes.txt)
$filePropertiesList | ForEach-Object {
    "File Name: $($_.'File Name')"
    "MD5: $($_.MD5)"
    "SHA-256: $($_.'SHA-256')"
    ""
} | Out-File -FilePath $outputFilePath

Write-Host "File properties exported to: $outputFilePath"
