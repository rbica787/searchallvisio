# ==========================
# EDIT THESE TWO VALUES ONLY
# ==========================

$FolderPath = "C:\Alerton\Compass\2.0\SYSERCO\CPOL2\ddc"
$SearchString = "2211"

# ==========================
# DO NOT EDIT BELOW THIS LINE
# ==========================

if (-not (Test-Path $FolderPath)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resultsFile = Join-Path $scriptDir "searchresults.txt"

if (Test-Path $resultsFile) {
    Add-Content -Path $resultsFile -Value ""
    Add-Content -Path $resultsFile -Value ""
} else {
    New-Item -Path $resultsFile -ItemType File -Force | Out-Null
}

Add-Content -Path $resultsFile -Value "Search run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Content -Path $resultsFile -Value "Search string: $SearchString"
Add-Content -Path $resultsFile -Value "Folder searched: $FolderPath"
Add-Content -Path $resultsFile -Value "----------------------------------------"

$files = Get-ChildItem -Path $FolderPath -File -Recurse -Include *.vsd, *.vsdx -ErrorAction SilentlyContinue

if (-not $files) {
    Write-Host "No Visio files found."
    exit 0
}

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $true
$visio.AlertResponse = 7

function Wait-ForVisioDocument {
    param($Document)

    $attempts = 0

    while ($attempts -lt 80) {
        try {
            if ($Document.Pages.Count -gt 0) {
                Start-Sleep -Milliseconds 500
                return $true
            }
        } catch {}

        Start-Sleep -Milliseconds 500
        $attempts++
    }

    return $false
}

function Test-ShapeForString {
    param(
        $Shape,
        [string]$SearchString
    )

    # Search visible shape text
    try {
        if ($Shape.Text -and $Shape.Text.ToLower().Contains($SearchString.ToLower())) {
            return $true
        }
    } catch {}

    # Search shape name
    try {
        if ($Shape.Name -and $Shape.Name.ToLower().Contains($SearchString.ToLower())) {
            return $true
        }
    } catch {}

    # Search shape name used in Visio UI
    try {
        if ($Shape.NameU -and $Shape.NameU.ToLower().Contains($SearchString.ToLower())) {
            return $true
        }
    } catch {}

    # Search Shape Data / Custom Properties
    try {
        $sectionExists = $Shape.SectionExists(243, 0)

        if ($sectionExists -ne 0) {
            $rowCount = $Shape.RowCount(243)

            for ($i = 0; $i -lt $rowCount; $i++) {
                try {
                    $label = $Shape.CellsSRC(243, $i, 2).ResultStr("")
                    $value = $Shape.CellsSRC(243, $i, 0).ResultStr("")

                    if ($label -and $label.ToLower().Contains($SearchString.ToLower())) {
                        return $true
                    }

                    if ($value -and $value.ToLower().Contains($SearchString.ToLower())) {
                        return $true
                    }
                } catch {}
            }
        }
    } catch {}

    # Search grouped shapes recursively
    try {
        foreach ($subShape in $Shape.Shapes) {
            if (Test-ShapeForString -Shape $subShape -SearchString $SearchString) {
                return $true
            }
        }
    } catch {}

    return $false
}

function Test-DocumentForString {
    param(
        $Document,
        [string]$SearchString
    )

    foreach ($page in $Document.Pages) {
        foreach ($shape in $page.Shapes) {
            if (Test-ShapeForString -Shape $shape -SearchString $SearchString) {
                return $true
            }
        }
    }

    return $false
}

function Close-VisioDocument {
    param($Document)

    try {
        $Document.Saved = $true
    } catch {}

    try {
        $Document.Close()
    } catch {}
}

foreach ($file in $files) {

    Write-Host "Opening: $($file.FullName)"

    try {
        $doc = $visio.Documents.Open($file.FullName)
    } catch {
        Write-Warning "Could not open: $($file.FullName)"
        continue
    }

    $loaded = Wait-ForVisioDocument -Document $doc

    if (-not $loaded) {
        Write-Warning "File did not fully load: $($file.Name)"
        Close-VisioDocument -Document $doc
        continue
    }

    $found = Test-DocumentForString -Document $doc -SearchString $SearchString

    if ($found) {
        Write-Host "FOUND in: $($file.Name)"
        Add-Content -Path $resultsFile -Value $file.Name

        # Leave matching Visio document open
    }
    else {
        Write-Host "Not found in: $($file.Name). Closing file."
        Close-VisioDocument -Document $doc
    }
}

Write-Host "Search complete."
Write-Host "Results saved to: $resultsFile"