<#
.SYNOPSIS
    This script downloads images from URLs in .eml files, renames the files, and organizes them into subfolders.

.DESCRIPTION
    This script assumes that the user has a folder of .eml files containing the emails from LineLeader, a service usually used by Lightbridge Academy. 
    The script renames the files to YYYY-MM-DD.eml format, generates a folder in YYYY-MM-DD format, downloads all linked assets referenced in each .eml file, and deletes .pdf .png, .html, and .css files. These files are typically not needed for the expected output.
    The script also deletes files with no extension and detects and deletes files with duplicate contents, keeping the older file. Finally, the script renumbers the files starting at 1, and changes its creation date to 5PM of the date of the email in 1 minute increments.

.PARAMETER FolderPath
    This parameter is the path to the folder containing .eml files.

.EXAMPLE
    Place .eml files in a folder.
    Run the script with the folder path as the parameter.
    The script will create subfolders for each .eml file, download linked assets, and organize the files into the subfolders.

.VERSION
    1.0

.RELEASEDATE
    2024-05-03

.AUTHOR
    Andrew Branagan
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath
)

# Get all files in the folder
$files = Get-ChildItem -Path $FolderPath -File

<## Set the path to the folder containing the files
$folderPath = "C:\Users\abranagan\ProjectLBA\2024-04-02 LBA Catchup"#>

# Get all files in the folder
$files = Get-ChildItem -Path $folderPath -File

foreach ($file in $files) {
    # Extract the date from the filename using regex
    $date = $file.BaseName -match '\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b'

    if ($date) {
        # Parse the extracted date into YYYY-MM-DD format
        $newName = Get-Date $Matches[0] -Format 'yyyy-MM-dd'

        # Construct the new filename with the date in YYYY-MM-DD format
        $newFileName = $newName + $file.Extension

        # Rename the file
        Rename-Item -Path $file.FullName -NewName $newFileName -Force
    }
    else {
        Write-Host "No date found in $($file.Name). Skipping..."
    }
}

$files = Get-ChildItem -Path $folderPath -File
# Create a subfolder for each .eml file
foreach ($file in $files) {
    if ($file.Extension -eq ".eml") {
        $subfolderName = $file.BaseName
        $subfolderPath = Join-Path -Path $folderPath -ChildPath $subfolderName
        New-Item -ItemType Directory -Path $subfolderPath -Force
    }
}

$files = Get-ChildItem -Path $folderPath -Filter "*.eml" -File -Recurse

foreach ($file in $files) {
    $subfolderName = $file.BaseName
    $subfolderPath = Join-Path -Path $folderPath -ChildPath $subfolderName

    $content = Get-Content -Raw $file.FullName | ForEach-Object {
        $_ -replace '=\r?\n', '' -replace '&amp;', '&' -replace '=3D', '='
    }

    $urls = $content | Select-String -Pattern 'https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)' -AllMatches | ForEach-Object {
        if ($_ -notmatch '\.(png|jpg)$') {
            $_.Matches.Value
        }
    }

    $counter = 1
    foreach ($url in $urls) {
        $response = Invoke-WebRequest -Uri $url -Method Head
        $contentType = $response.Headers["Content-Type"]
        $extension = $contentType.Split("/")[-1]
        $outputPath = Join-Path -Path $subfolderPath -ChildPath ($subfolderName + "_" + $counter.ToString("00") + "." + $extension)
        Invoke-WebRequest -Uri $url -OutFile $outputPath
        $counter++

        # Delete files with names containing "html" or "css"
        $filesToDelete = $subfolderPath | Get-ChildItem -File | Where-Object { $_.Name -match 'html|css|png|pdf' }
        $filesToDelete | Remove-Item -Force

        # Delete files with no extension
        $filesToDelete = $subfolderPath | Get-ChildItem -File | Where-Object { $_.Extension -eq '' }
        $filesToDelete | Remove-Item -Force

        # Detect files with duplicate contents and delete the younger one
        $filesToDelete = Get-ChildItem -Path $subfolderPath -File | Group-Object -Property @{Expression={Get-Content -Raw $_.FullName}} | Where-Object {$_.Count -gt 1} | ForEach-Object {
            $_.Group | Sort-Object -Property LastWriteTime | Select-Object -Skip 1
        }
        $filesToDelete | Remove-Item -Force

        # Renumber the files starting at 1
        $renumberedFiles = Get-ChildItem -Path $subfolderPath -File | Sort-Object LastWriteTime
        $counter = 1
        foreach ($file in $renumberedFiles) {
            $newFileName = "{0}_{1:D2}{2}" -f $subfolderName, $counter, $file.Extension
            $newFilePath = Join-Path -Path $subfolderPath -ChildPath $newFileName
            Rename-Item -Path $file.FullName -NewName $newFileName -Force
            $counter++
        }

    }

}

$folders = Get-ChildItem -Path $folderPath -Directory -Recurse

foreach ($folder in $folders) {
    $subfolderPath = $folder.FullName
    $files = Get-ChildItem -Path $subfolderPath -Filter "*.*" -File

    $dateTaken = Get-Date -Year $folder.Name.Substring(0, 4) -Month $folder.Name.Substring(5, 2) -Day $folder.Name.Substring(8, 2) -Hour 17 -Minute 0 -Second 0

    $order = 0
    foreach ($file in $files) {
        $order++
        $newDateTaken = $dateTaken.AddMinutes($order)
        Set-ItemProperty -Path $file.FullName -Name "CreationTime" -Value $newDateTaken
        Set-ItemProperty -Path $file.FullName -Name "LastWriteTime" -Value $newDateTaken
    }
}
