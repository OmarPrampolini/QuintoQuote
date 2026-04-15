param(
    [string]$TargetDir = (Join-Path $HOME "QuintoQuote"),
    [string]$Branch = "main",
    [string]$RepoUrl = "https://github.com/OmarPrampolini/QuintoQuote.git",
    [string]$ArchiveUrl = "https://github.com/OmarPrampolini/QuintoQuote/archive/refs/heads/main.zip",
    [string]$SourcePath = "",
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Resolve-PythonCommand {
    $candidates = @(
        @{ Path = "python"; Args = @() },
        @{ Path = "python3"; Args = @() },
        @{ Path = "py"; Args = @("-3") },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Python\\bin\\python3.exe"); Args = @() },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python314\\python.exe"); Args = @() },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python313\\python.exe"); Args = @() },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python312\\python.exe"); Args = @() },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python311\\python.exe"); Args = @() },
        @{ Path = (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python310\\python.exe"); Args = @() }
    )
    foreach ($candidate in $candidates) {
        try {
            & $candidate.Path @($candidate.Args + @("--version")) *> $null
            if ($LASTEXITCODE -eq 0) {
                return [pscustomobject]$candidate
            }
        } catch {
        }
    }
    throw "Python 3.10+ non trovato. Installa Python e riprova."
}

function Invoke-Python($PythonCommand, [string[]]$Arguments) {
    & $PythonCommand.Path @($PythonCommand.Args + $Arguments)
    if ($LASTEXITCODE -ne 0) {
        throw "Comando Python fallito: $($Arguments -join ' ')"
    }
}

function Sync-LocalSource([string]$FromPath, [string]$ToPath) {
    $source = (Resolve-Path $FromPath).Path
    New-Item -ItemType Directory -Force -Path $ToPath | Out-Null
    $exclude = @(".git", ".venv", "__pycache__", ".quintoquote_tmp", "output_preventivi", "quintoquote.egg-info")
    Get-ChildItem -LiteralPath $source -Force |
        Where-Object { $exclude -notcontains $_.Name } |
        ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination $ToPath -Recurse -Force
        }
}

function Ensure-Repo([string]$Destination, [string]$BranchName, [string]$GitUrl, [string]$ZipUrl, [string]$LocalSource) {
    if ($LocalSource) {
        Write-Step "Copio il repository locale in $Destination"
        Sync-LocalSource -FromPath $LocalSource -ToPath $Destination
        return
    }

    if (Test-Path (Join-Path $Destination ".git")) {
        $git = Get-Command git -ErrorAction SilentlyContinue
        if ($git) {
            Write-Step "Aggiorno il repository esistente"
            & $git.Source -C $Destination pull --ff-only origin $BranchName
            if ($LASTEXITCODE -ne 0) {
                throw "Aggiornamento git fallito."
            }
            return
        }
    }

    if (Test-Path (Join-Path $Destination "pyproject.toml")) {
        Write-Step "Uso la cartella esistente in $Destination"
        return
    }

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git) {
        Write-Step "Clono il repository GitHub"
        & $git.Source clone --depth 1 --branch $BranchName $GitUrl $Destination
        if ($LASTEXITCODE -eq 0) {
            return
        }
    }

    Write-Step "Scarico l'archivio GitHub"
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    $runtimeTmp = Join-Path $Destination ".quintoquote_tmp"
    $zipPath = Join-Path $runtimeTmp "repo.zip"
    $extractRoot = Join-Path $runtimeTmp "repo"
    New-Item -ItemType Directory -Force -Path $runtimeTmp | Out-Null
    if (Test-Path $extractRoot) {
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -UseBasicParsing
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
    $expanded = Get-ChildItem -LiteralPath $extractRoot -Directory | Select-Object -First 1
    if (-not $expanded) {
        throw "Archivio GitHub non valido."
    }
    Get-ChildItem -LiteralPath $expanded.FullName -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
    }
}

function Ensure-Tesseract {
    Write-Step "Controllo Tesseract OCR"
    $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($tesseractCmd) {
        Write-Host "Tesseract OCR già installato." -ForegroundColor Green
        return $tesseractCmd.Source
    }

    $commonPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    if (Test-Path $commonPath) {
        Write-Host "Tesseract OCR trovato in $commonPath." -ForegroundColor Green
        return $commonPath
    }

    Write-Step "Tesseract OCR non trovato. Tento l'installazione tramite winget..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Tesseract OCR è obbligatorio ma winget non è disponibile. Installa Tesseract manualmente da https://github.com/UB-Mannheim/tesseract/wiki e riprova."
    }

    & $winget.Source install --id UB_Mannheim.TesseractOCR --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "Installazione di Tesseract OCR fallita. Installa Tesseract manualmente per continuare."
    }

    # Refresh path or check common location again
    if (Test-Path $commonPath) {
        return $commonPath
    }
    
    $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($tesseractCmd) {
        return $tesseractCmd.Source
    }

    throw "Tesseract OCR installato ma non trovato nel PATH. Riavvia la shell o aggiungi manualmente il percorso di installazione al PATH."
}

$python = Resolve-PythonCommand
$tesseractPath = Ensure-Tesseract
Ensure-Repo -Destination $TargetDir -BranchName $Branch -GitUrl $RepoUrl -ZipUrl $ArchiveUrl -LocalSource $SourcePath

$repoRoot = (Resolve-Path $TargetDir).Path
$runtimeTmp = Join-Path $repoRoot ".quintoquote_tmp"
New-Item -ItemType Directory -Force -Path $runtimeTmp | Out-Null
$env:TEMP = $runtimeTmp
$env:TMP = $runtimeTmp

Write-Step "Creo o aggiorno la virtualenv"
$venvPath = Join-Path $repoRoot ".venv"
$venvPython = Join-Path $venvPath "Scripts\\python.exe"
try {
    Invoke-Python $python @("-m", "venv", $venvPath)
} catch {
    if (-not (Test-Path $venvPython)) {
        throw
    }
    Write-Host "Bootstrap venv incompleto, continuo con il recupero di pip dentro la virtualenv." -ForegroundColor Yellow
}

Write-Step "Aggiorno pip nella virtualenv"
Invoke-Python $python @("-m", "pip", "--python", $venvPath, "install", "--upgrade", "pip")

Write-Step "Installo QuintoQuote"
Push-Location $repoRoot
try {
    Invoke-Python $python @("-m", "pip", "--python", $venvPath, "install", "-e", ".")
} finally {
    Pop-Location
}

$launcher = Join-Path $repoRoot ".venv\\Scripts\\quintoquote.exe"
if (-not (Test-Path $launcher)) {
    throw "Launcher non trovato: $launcher"
}

Write-Host ""
Write-Host "QuintoQuote installato in $repoRoot" -ForegroundColor Green
Write-Host "OCR locale attivo: $tesseractPath" -ForegroundColor Green
Write-Host "Avvio successivo: $launcher start" -ForegroundColor DarkGray
Write-Host ""

if (-not $NoLaunch) {
    Write-Step "Avvio QuintoQuote"
    & $launcher start
}
