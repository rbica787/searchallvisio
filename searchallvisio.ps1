# ==========================
# EDIT THESE VALUES ONLY
# ==========================

$FolderPath = "C:\Alerton\Compass\2.0\SYSERCO\BACETH\ddc"
$SearchString = "hello"

# Set to $true to search subfolders
# Set to $false to search only this folder
$RecursiveSearch = $true

# Save extracted XML copies for files where a match is found
$SaveMatchingXml = $true

# ==========================
# DO NOT EDIT BELOW THIS LINE
# ==========================

if (-not (Test-Path -LiteralPath $FolderPath)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

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
Add-Content -LiteralPath $resultsFile "Search method: Direct .vsdx XML parsing"
Add-Content -LiteralPath $resultsFile "----------------------------------------"

if ($RecursiveSearch) {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -Recurse -Filter "*.vsdx" -ErrorAction SilentlyContinue
} else {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -Filter "*.vsdx" -ErrorAction SilentlyContinue
}

$filesScanned = 0
$matchesFound = 0

function Convert-XmlToSearchText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Text.StringBuilder]$Builder
    )

    if ($null -eq $Node) {
        return
    }

    if ($Node.Value) {
        [void]$Builder.AppendLine($Node.Value)
    }

    if ($Node.Attributes) {
        foreach ($attr in $Node.Attributes) {
            if ($attr.Name) {
                [void]$Builder.AppendLine($attr.Name)
            }

            if ($attr.Value) {
                [void]$Builder.AppendLine($attr.Value)
            }
        }
    }

    foreach ($child in $Node.ChildNodes) {
        Convert-XmlToSearchText -Node $child -Builder $Builder
    }
}

function Test-XmlTextForSearchString {
    param(
        [string]$XmlText,
        [string]$SearchString
    )

    $needle = $SearchString.ToLower()

    # Fast raw XML check first
    if ($XmlText.ToLower().Contains($needle)) {
        return $true
    }

    # Parsed XML check for text nodes and attributes
    try {
        [xml]$xml = $XmlText

        $builder = New-Object System.Text.StringBuilder
        Convert-XmlToSearchText -Node $xml -Builder $builder

        $searchBlob = $builder.ToString().ToLower()

        if ($searchBlob.Contains($needle)) {
            return $true
        }
    } catch {
        # If XML parsing fails, raw text check already happened.
    }

    return $false
}

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

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)

        foreach ($entry in $archive.Entries) {

            # Only inspect XML files inside the .vsdx package
            if (-not $entry.FullName.ToLower().EndsWith(".xml")) {
                continue
            }

            # These are the most relevant Visio XML areas:
            # visio/pages/page*.xml     = page shapes and text
            # visio/masters/master*.xml = master shapes
            # visio/document.xml        = document-level data
            # docProps/*.xml            = file properties
            $entryNameLower = $entry.FullName.ToLower()

            $isRelevantXml =
                $entryNameLower.StartsWith("visio/pages/") -or
                $entryNameLower.StartsWith("visio/masters/") -or
                $entryNameLower -eq "visio/document.xml" -or
                $entryNameLower.StartsWith("docprops/")

            if (-not $isRelevantXml) {
                continue
            }

            try {
                $stream = $entry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                $xmlText = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()

                if (Test-XmlTextForSearchString -XmlText $xmlText -SearchString $SearchString) {
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

    Write-Host "Scanning XML: $($file.FullName)"

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
            Add-Content -LiteralPath $resultsFile "    XML match: $entry"
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

Write-Host "Search complete."
Write-Host "Files scanned: $filesScanned"
Write-Host "Matches found: $matchesFound"
Write-Host "Results saved to: $resultsFile"

if ($SaveMatchingXml) {
    Write-Host "Matching XML saved to: $xmlOutputRoot"
}
