Add-Type -AssemblyName System.IO.Compression.FileSystem

$SaveMatchingXml = $true

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resultsFile = Join-Path $scriptDir "searchresults.txt"
$xmlOutputRoot = Join-Path $scriptDir "ExtractedXmlMatches"

if ($SaveMatchingXml -and -not (Test-Path -LiteralPath $xmlOutputRoot)) {
    New-Item -Path $xmlOutputRoot -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $resultsFile)) {
    New-Item -Path $resultsFile -ItemType File -Force | Out-Null
}

function Get-SearchSettings {
    do {
        $folderPath = Read-Host "Enter the folder path to search"

        if (-not (Test-Path -LiteralPath $folderPath)) {
            Write-Host "Folder not found: $folderPath"
            Write-Host ""
        }
    } while (-not (Test-Path -LiteralPath $folderPath))

    do {
        $recursiveAnswer = Read-Host "Search subfolders recursively? Enter Y or N"
        $recursiveAnswer = $recursiveAnswer.Trim().ToUpper()
    } while ($recursiveAnswer -ne "Y" -and $recursiveAnswer -ne "N")

    $recursiveSearch = ($recursiveAnswer -eq "Y")

    if ($recursiveSearch) {
        $files = Get-ChildItem -LiteralPath $folderPath -File -Recurse -Filter "*.vsdx" -ErrorAction SilentlyContinue
    } else {
        $files = Get-ChildItem -LiteralPath $folderPath -File -Filter "*.vsdx" -ErrorAction SilentlyContinue
    }

    if (-not $files) {
        Write-Host ""
        Write-Host "No .vsdx files found in this location."
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "File list loaded."
        Write-Host "VSDX files found: $($files.Count)"
        Write-Host ""
    }

    return @{
        FolderPath = $folderPath
        RecursiveSearch = $recursiveSearch
        Files = $files
    }
}

function Save-MatchingXmlEntry {
    param(
        [System.IO.Compression.ZipArchiveEntry]$Entry,
        [string]$XmlText,
        [string]$DrawingName,
        [string]$OutputRoot,
        [string]$SearchString
    )

    $safeDrawingName = [System.IO.Path]::GetFileNameWithoutExtension($DrawingName)
    $safeDrawingName = $safeDrawingName -replace '[\\/:*?"<>|]', '_'

    $safeSearchString = $SearchString -replace '[\\/:*?"<>|]', '_'

    $drawingFolder = Join-Path $OutputRoot $safeSearchString
    $drawingFolder = Join-Path $drawingFolder $safeDrawingName

    if (-not (Test-Path -LiteralPath $drawingFolder)) {
        New-Item -Path $drawingFolder -ItemType Directory -Force | Out-Null
    }

    $entryName = $Entry.FullName -replace '/', '_'
    $entryName = $entryName -replace '[\\:*?"<>|]', '_'

    $xmlPath = Join-Path $drawingFolder $entryName

    Set-Content -LiteralPath $xmlPath -Value $XmlText -Encoding UTF8
}

function Test-VsdxForSearchString {
    param(
        [System.IO.FileInfo]$File,
        [string]$SearchString,
        [bool]$SaveMatchingXml,
        [string]$XmlOutputRoot
    )

    $archive = $null
    $matchedEntries = @()
    $needle = $SearchString.ToLower()

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)

        foreach ($entry in $archive.Entries) {
            $entryNameLower = $entry.FullName.ToLower()

            # ONLY search actual drawing pages.
            # Master pages are intentionally ignored.
            if ($entryNameLower -notmatch '^visio/pages/page[0-9]+\.xml$') {
                continue
            }

            try {
                $stream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $xmlText = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()

                if ($xmlText.ToLower().Contains($needle)) {
                    $matchedEntries += $entry.FullName

                    if ($SaveMatchingXml) {
                        Save-MatchingXmlEntry `
                            -Entry $entry `
                            -XmlText $xmlText `
                            -DrawingName $File.Name `
                            -OutputRoot $XmlOutputRoot `
                            -SearchString $SearchString
                    }
                }
            } catch {}
        }
    } catch {
        return @{
            Error = $true
            Found = $false
            Entries = @()
        }
    } finally {
        if ($archive) {
            $archive.Dispose()
        }
    }

    return @{
        Error = $false
        Found = ($matchedEntries.Count -gt 0)
        Entries = $matchedEntries
    }
}

function Invoke-VsdxSearch {
    param(
        [string]$FolderPath,
        [bool]$RecursiveSearch,
        [array]$Files,
        [string]$SearchString
    )

    $filesScanned = 0
    $matchesFound = 0

    Add-Content -LiteralPath $resultsFile ""
    Add-Content -LiteralPath $resultsFile ""
    Add-Content -LiteralPath $resultsFile "Search run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-Content -LiteralPath $resultsFile "Search string: $SearchString"
    Add-Content -LiteralPath $resultsFile "Folder searched: $FolderPath"
    Add-Content -LiteralPath $resultsFile "Recursive search: $RecursiveSearch"
    Add-Content -LiteralPath $resultsFile "Search method: Direct .vsdx page XML parsing only"
    Add-Content -LiteralPath $resultsFile "Master pages ignored: True"
    Add-Content -LiteralPath $resultsFile "----------------------------------------"

    foreach ($file in $Files) {
        $filesScanned++

        Write-Host "Scanning page XML only: $($file.FullName)"

        $result = Test-VsdxForSearchString `
            -File $file `
            -SearchString $SearchString `
            -SaveMatchingXml $SaveMatchingXml `
            -XmlOutputRoot $xmlOutputRoot

        if ($result.Error) {
            Write-Warning "Could not read as .vsdx package: $($file.FullName)"
            Add-Content -LiteralPath $resultsFile "Could not read: $($file.Name)"
            continue
        }

        if ($result.Found) {
            $matchesFound++

            Write-Host "FOUND in: $($file.Name)"
            Add-Content -LiteralPath $resultsFile $file.Name

            foreach ($entry in $result.Entries) {
                Add-Content -LiteralPath $resultsFile "    Page XML match: $entry"
            }
        } else {
            Write-Host "Not found in: $($file.Name)."
        }
    }

    Add-Content -LiteralPath $resultsFile "----------------------------------------"
    Add-Content -LiteralPath $resultsFile "Files scanned: $filesScanned"
    Add-Content -LiteralPath $resultsFile "Matches found: $matchesFound"

    if ($SaveMatchingXml) {
        Add-Content -LiteralPath $resultsFile "Extracted matching XML folder: $xmlOutputRoot"
    }

    Add-Content -LiteralPath $resultsFile "Search complete."

    Write-Host ""
    Write-Host "Search complete."
    Write-Host "Files scanned: $filesScanned"
    Write-Host "Matches found: $matchesFound"
    Write-Host "Results saved to: $resultsFile"

    if ($SaveMatchingXml) {
        Write-Host "Matching page XML saved to: $xmlOutputRoot"
    }

    Write-Host ""
}

$settings = Get-SearchSettings

while (-not $settings.Files -or $settings.Files.Count -eq 0) {
    do {
        $retryChoice = Read-Host "Would you like to choose a different folder/settings? Enter Y or N"
        $retryChoice = $retryChoice.Trim().ToUpper()
    } while ($retryChoice -ne "Y" -and $retryChoice -ne "N")

    if ($retryChoice -eq "Y") {
        $settings = Get-SearchSettings
    } else {
        Write-Host "Program finished."
        exit 0
    }
}

$FolderPath = $settings.FolderPath
$RecursiveSearch = $settings.RecursiveSearch
$files = $settings.Files

do {
    do {
        $SearchString = Read-Host "Enter the search string"

        if ([string]::IsNullOrWhiteSpace($SearchString)) {
            Write-Host "Search string cannot be blank."
            Write-Host ""
        }
    } while ([string]::IsNullOrWhiteSpace($SearchString))

    Invoke-VsdxSearch `
        -FolderPath $FolderPath `
        -RecursiveSearch $RecursiveSearch `
        -Files $files `
        -SearchString $SearchString

    do {
        Write-Host "What would you like to do next?"
        Write-Host "[S] Search another string using the same folder/settings"
        Write-Host "[C] Change search folder/settings"
        Write-Host "[E] Exit"

        $nextChoice = Read-Host "Enter S, C, or E"
        $nextChoice = $nextChoice.Trim().ToUpper()
    } while ($nextChoice -ne "S" -and $nextChoice -ne "C" -and $nextChoice -ne "E")

    if ($nextChoice -eq "C") {
        $settings = Get-SearchSettings

        while (-not $settings.Files -or $settings.Files.Count -eq 0) {
            do {
                $retryChoice = Read-Host "Would you like to choose a different folder/settings? Enter Y or N"
                $retryChoice = $retryChoice.Trim().ToUpper()
            } while ($retryChoice -ne "Y" -and $retryChoice -ne "N")

            if ($retryChoice -eq "Y") {
                $settings = Get-SearchSettings
            } else {
                Write-Host "Program finished."
                exit 0
            }
        }

        $FolderPath = $settings.FolderPath
        $RecursiveSearch = $settings.RecursiveSearch
        $files = $settings.Files
    }

} while ($nextChoice -ne "E")

Write-Host ""
Write-Host "Program finished."
