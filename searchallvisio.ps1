# ==========================
# VSDX XML SEARCH SCRIPT
# Searches page XML only
# Master pages are ignored
# ==========================

Add-Type -AssemblyName System.IO.Compression.FileSystem

$FolderPath = Read-Host "Enter the folder path to search"
$SearchString = Read-Host "Enter the search string"

do {
    $recursiveAnswer = Read-Host "Search subfolders recursively? Enter Y or N"
    $recursiveAnswer = $recursiveAnswer.Trim().ToUpper()
} while ($recursiveAnswer -ne "Y" -and $recursiveAnswer -ne "N")

$RecursiveSearch = ($recursiveAnswer -eq "Y")

$SaveMatchingXml = $true

if (-not (Test-Path -LiteralPath $FolderPath)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($SearchString)) {
    Write-Error "Search string cannot be blank."
    exit 1
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resultsFile = Join-Path $scriptDir "searchresults.txt"
$xmlOutputRoot = Join-Path $scriptDir "ExtractedXmlMatches"

if ($SaveMatchingXml -and -not (Test-Path -LiteralPath $xmlOutputRoot)) {
    New-Item -Path $xmlOutputRoot -ItemType Directory -Force | Out-Null
}

if (Test-Path -LiteralPath $resultsFile) {
    Add-Content -LiteralPath $resultsFile ""
    Add-Content -LiteralPath $resultsFile ""
} else {
    New-Item -Path $resultsFile -ItemType File -Force | Out-Null
}

Add-Content -LiteralPath $resultsFile "Search run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Content -LiteralPath $resultsFile "Search string: $SearchString"
Add-Content -LiteralPath $resultsFile "Folder searched: $FolderPath"
Add-Content -LiteralPath $resultsFile "Recursive search: $RecursiveSearch"
Add-Content -LiteralPath $resultsFile "Search method: Direct .vsdx page XML parsing only"
Add-Content -LiteralPath $resultsFile "Master pages ignored: True"
Add-Content -LiteralPath $resultsFile "----------------------------------------"

if ($RecursiveSearch) {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -Recurse -Filter "*.vsdx" -ErrorAction SilentlyContinue
} else {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -Filter "*.vsdx" -ErrorAction SilentlyContinue
}

$filesScanned = 0
$matchesFound = 0

function Save-MatchingXmlEntry {
    param(
        [System.IO.Compression.ZipArchiveEntry]$Entry,
        [string]$XmlText,
        [string]$DrawingName,
        [string]$OutputRoot
    )

    $safeDrawingName = [System.IO.Path]::GetFileNameWithoutExtension($DrawingName)
    $safeDrawingName = $safeDrawingName -replace '[\\/:*?"<>|]', '_'

    $drawingFolder = Join-Path $OutputRoot $safeDrawingName

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
                            -OutputRoot $XmlOutputRoot
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

if (-not $files) {
    Write-Host "No .vsdx files found."

    Add-Content -LiteralPath $resultsFile "No .vsdx files found."
    Add-Content -LiteralPath $resultsFile "----------------------------------------"
    Add-Content -LiteralPath $resultsFile "Files scanned: 0"
    Add-Content -LiteralPath $resultsFile "Matches found: 0"
    Add-Content -LiteralPath $resultsFile "Search complete."

    exit 0
}

foreach ($file in $files) {
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
