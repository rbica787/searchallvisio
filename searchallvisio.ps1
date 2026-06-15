# ==========================
# EDIT THESE THREE VALUES ONLY
# ==========================

$FolderPath = "C:\Alerton\Compass\2.0\SYSERCO\BACETH\ddc"
$SearchString = "hello"

# Set to $true to search subfolders
# Set to $false to search only this folder
$RecursiveSearch = $true

# ==========================
# DO NOT EDIT BELOW THIS LINE
# ==========================

if (-not (Test-Path -LiteralPath $FolderPath)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$resultsFile = Join-Path $scriptDir "searchresults.txt"
$watchdogFile = Join-Path $env:TEMP "VisualLogicPromptWatchdog.ps1"

@'
Add-Type -AssemblyName System.Windows.Forms

while ($true) {
    try {
        $shell = New-Object -ComObject WScript.Shell

        $titles = @(
            "Select Controller for Database Creation (BD1)",
            "Select Controller for Database Creation (BD2)",
            "Select Controller for Database Creation (BD3)",
            "Select Controller for Database Creation (BD4)",
            "Select Controller for Database Creation (BD5)",
            "Select Controller for Database Creation (BD6)",
            "Select Controller for Database Creation (BD7)",
            "Select Controller for Database Creation (BD8)",
            "Select Controller for Database Creation (BD9)"
        )

        foreach ($title in $titles) {
            if ($shell.AppActivate($title)) {
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait("{HOME}")
                Start-Sleep -Milliseconds 200
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Start-Sleep -Milliseconds 1000
                break
            }
        }
    } catch {}

    Start-Sleep -Milliseconds 300
}
'@ | Set-Content -LiteralPath $watchdogFile -Encoding UTF8

$watchdog = Start-Process powershell.exe `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogFile`"" `
    -PassThru

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
Add-Content -LiteralPath $resultsFile "----------------------------------------"

if ($RecursiveSearch) {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".vsd", ".vsdx" }
} else {
    $files = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".vsd", ".vsdx" }
}

$filesScanned = 0
$matchesFound = 0

if (-not $files) {
    Write-Host "No Visio files found."

    Add-Content -LiteralPath $resultsFile "No Visio files found."
    Add-Content -LiteralPath $resultsFile "----------------------------------------"
    Add-Content -LiteralPath $resultsFile "Files scanned: 0"
    Add-Content -LiteralPath $resultsFile "Matches found: 0"
    Add-Content -LiteralPath $resultsFile "Search complete."

    if ($watchdog -and -not $watchdog.HasExited) {
        Stop-Process -Id $watchdog.Id -Force
    }

    try {
        if ($watchdogFile -and (Test-Path -LiteralPath $watchdogFile)) {
            Remove-Item -LiteralPath $watchdogFile -Force -ErrorAction SilentlyContinue
        }
    } catch {}

    exit 0
}

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $true
$visio.AlertResponse = 7

function Wait-ForVisioDocument {
    param($Document)

    $attempts = 0

    while ($attempts -lt 180) {
        try {
            if ($Document.Pages.Count -gt 0) {
                Start-Sleep -Milliseconds 700
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

    $needle = $SearchString.ToLower()

    try {
        if ($Shape.Text -and $Shape.Text.ToLower().Contains($needle)) {
            return $true
        }
    } catch {}

    try {
        if ($Shape.Name -and $Shape.Name.ToLower().Contains($needle)) {
            return $true
        }
    } catch {}

    try {
        if ($Shape.NameU -and $Shape.NameU.ToLower().Contains($needle)) {
            return $true
        }
    } catch {}

    try {
        $sectionExists = $Shape.SectionExists(243, 0)

        if ($sectionExists -ne 0) {
            $rowCount = $Shape.RowCount(243)

            for ($i = 0; $i -lt $rowCount; $i++) {
                try {
                    $label = $Shape.CellsSRC(243, $i, 2).ResultStr("")
                    $value = $Shape.CellsSRC(243, $i, 0).ResultStr("")

                    if ($label -and $label.ToLower().Contains($needle)) {
                        return $true
                    }

                    if ($value -and $value.ToLower().Contains($needle)) {
                        return $true
                    }
                } catch {}
            }
        }
    } catch {}

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
        $Document.Close()
        Start-Sleep -Milliseconds 700
    } catch {
        Write-Warning "Could not close document."
    }
}

try {
    foreach ($file in $files) {

        $filesScanned++

        Write-Host "Opening: $($file.FullName)"

        try {
            $doc = $visio.Documents.Open($file.FullName)
        } catch {
            Write-Warning "Could not open: $($file.FullName)"
            Add-Content -LiteralPath $resultsFile "Could not open: $($file.Name)"
            continue
        }

        $loaded = Wait-ForVisioDocument -Document $doc

        if (-not $loaded) {
            Write-Warning "File did not fully load: $($file.Name)"
            Add-Content -LiteralPath $resultsFile "File did not fully load: $($file.Name)"
            Close-VisioDocument -Document $doc
            continue
        }

        $found = Test-DocumentForString -Document $doc -SearchString $SearchString

        if ($found) {
            $matchesFound++
            Write-Host "FOUND in: $($file.Name)"
            Add-Content -LiteralPath $resultsFile $file.Name
        }
        else {
            Write-Host "Not found in: $($file.Name). Closing file."
            Close-VisioDocument -Document $doc
        }

        Start-Sleep -Milliseconds 500
    }
}
finally {
    Add-Content -LiteralPath $resultsFile "----------------------------------------"
    Add-Content -LiteralPath $resultsFile "Files scanned: $filesScanned"
    Add-Content -LiteralPath $resultsFile "Matches found: $matchesFound"
    Add-Content -LiteralPath $resultsFile "Search complete."

    if ($watchdog -and -not $watchdog.HasExited) {
        Stop-Process -Id $watchdog.Id -Force
    }

    try {
        if ($watchdogFile -and (Test-Path -LiteralPath $watchdogFile)) {
            Remove-Item -LiteralPath $watchdogFile -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

Write-Host "Search complete."
Write-Host "Files scanned: $filesScanned"
Write-Host "Matches found: $matchesFound"
Write-Host "Results saved to: $resultsFile"
