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
$watchdogFile = Join-Path $env:TEMP "VisualLogicPromptWatchdog.ps1"

# Create separate watchdog script to handle VisualLogic popup while main script is blocked
@'
Add-Type -AssemblyName System.Windows.Forms

while ($true) {
    try {
        $shell = New-Object -ComObject WScript.Shell

        if ($shell.AppActivate("Select Controller for Database Creation (BD3)")) {
            Start-Sleep -Milliseconds 300

            [System.Windows.Forms.SendKeys]::SendWait("{HOME}")
            Start-Sleep -Milliseconds 200

            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Milliseconds 1000
        }
    } catch {}

    Start-Sleep -Milliseconds 300
}
'@ | Set-Content -Path $watchdogFile -Encoding UTF8

$watchdog = Start-Process powershell.exe `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogFile`"" `
    -PassThru

if (Test-Path $resultsFile) {
    Add-Content $resultsFile ""
    Add-Content $resultsFile ""
} else {
    New-Item -Path $resultsFile -ItemType File -Force | Out-Null
}

Add-Content $resultsFile "Search run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Add-Content $resultsFile "Search string: $SearchString"
Add-Content $resultsFile "Folder searched: $FolderPath"
Add-Content $resultsFile "----------------------------------------"

$files = Get-ChildItem -Path $FolderPath -File -Recurse -Include *.vsd, *.vsdx -ErrorAction SilentlyContinue

if (-not $files) {
    Write-Host "No Visio files found."

    if ($watchdog -and -not $watchdog.HasExited) {
        Stop-Process -Id $watchdog.Id -Force
    }

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
        }
        else {
            Write-Host "Not found in: $($file.Name). Closing file."
            Close-VisioDocument -Document $doc
        }

        Start-Sleep -Milliseconds 500
    }
}
finally {
    if ($watchdog -and -not $watchdog.HasExited) {
        Stop-Process -Id $watchdog.Id -Force
    }

    if (Test-Path $watchdogFile) {
        Remove-Item $watchdogFile -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "Search complete."
Write-Host "Results saved to: $resultsFile"
