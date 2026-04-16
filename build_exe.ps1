param(
    [switch]$Clean,
    [switch]$SkipZip,
    [switch]$SkipOcrBundle
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Ensure-CommandSuccess([string]$Message) {
    if ($LASTEXITCODE -ne 0) {
        throw $Message
    }
}

function Resolve-TesseractRoot {
    $configured = $env:QUINTOQUOTE_TESSERACT_PATH
    $candidates = @()
    if ($configured) {
        $candidates += $configured
    }
    $common = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    if (Test-Path $common) {
        $candidates += $common
    }
    $cmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($cmd) {
        $candidates += $cmd.Source
    }
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return Split-Path -Parent $candidate
        }
    }
    return ""
}

function Assert-TrustedModelSource([string]$Url, [string]$Name) {
    $uri = [System.Uri]$Url
    if ($uri.Scheme -ne "https") {
        throw "Origine non valida per ${Name}: è richiesto HTTPS."
    }
    if ($uri.Host -ne "raw.githubusercontent.com") {
        throw "Origine non valida per ${Name}: host non autorizzato ($($uri.Host))."
    }
    $expectedPath = "/tesseract-ocr/tessdata_fast/main/$Name"
    if ($uri.AbsolutePath -ne $expectedPath) {
        throw "Origine non valida per ${Name}: percorso inatteso ($($uri.AbsolutePath))."
    }
}

function Install-TessdataModel([string]$Url, [string]$TargetFile, [string]$ExpectedSha256) {
    $name = Split-Path -Leaf $TargetFile
    Assert-TrustedModelSource -Url $Url -Name $name

    $tempFile = "$TargetFile.download"
    if (Test-Path $tempFile) {
        Remove-Item -LiteralPath $tempFile -Force
    }

    Invoke-WebRequest -Uri $Url -OutFile $tempFile -UseBasicParsing
    $actualSha256 = (Get-FileHash -LiteralPath $tempFile -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualSha256 -ne $ExpectedSha256.ToLowerInvariant()) {
        Remove-Item -LiteralPath $tempFile -Force
        throw "Hash SHA-256 non valida per $name. Attesa: $ExpectedSha256 - Ottenuta: $actualSha256"
    }

    Move-Item -LiteralPath $tempFile -Destination $TargetFile -Force
}

function Copy-TesseractBundle([string]$SourceRoot, [string]$DestinationRoot) {
    if (-not $SourceRoot) {
        throw "Tesseract OCR non trovato. Installa Tesseract oppure imposta QUINTOQUOTE_TESSERACT_PATH."
    }

    $dest = Join-Path $DestinationRoot "ocr"
    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    Copy-Item -LiteralPath (Join-Path $SourceRoot "tesseract.exe") -Destination $dest -Force
    Get-ChildItem -LiteralPath $SourceRoot -Filter "*.dll" | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
    }

    $sourceTessdata = Join-Path $SourceRoot "tessdata"
    if (-not (Test-Path $sourceTessdata)) {
        throw "Cartella tessdata non trovata in $SourceRoot"
    }

    $destTessdata = Join-Path $dest "tessdata"
    New-Item -ItemType Directory -Force -Path $destTessdata | Out-Null
    Get-ChildItem -LiteralPath $sourceTessdata | ForEach-Object {
        if ($_.PSIsContainer) {
            Copy-Item -LiteralPath $_.FullName -Destination $destTessdata -Recurse -Force
        } elseif ($_.Name -match '\.(traineddata|user-patterns|user-words)$') {
            Copy-Item -LiteralPath $_.FullName -Destination $destTessdata -Force
        }
    }

    $requiredModels = @(
        @{
            Name = "eng.traineddata"
            Url = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/eng.traineddata"
            Sha256 = "7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2"
        },
        @{
            Name = "ita.traineddata"
            Url = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/ita.traineddata"
            Sha256 = "b8f89e1e785118dac4d51ae042c029a64edb5c3ee42ef73027a6d412748d8827"
        },
        @{
            Name = "osd.traineddata"
            Url = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/osd.traineddata"
            Sha256 = "9cf5d576fcc47564f11265841e5ca839001e7e6f38ff7f7aacf46d15a96b00ff"
        }
    )
    foreach ($model in $requiredModels) {
        $targetFile = Join-Path $destTessdata $model.Name
        if (Test-Path $targetFile) {
            continue
        }
        Write-Step "Scarico e verifico $($model.Name) per il bundle OCR"
        Install-TessdataModel -Url $model.Url -TargetFile $targetFile -ExpectedSha256 $model.Sha256
    }

    if (-not (Test-Path (Join-Path $destTessdata "ita.traineddata"))) {
        throw "ita.traineddata non disponibile nel bundle OCR."
    }
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $repoRoot

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Virtualenv non trovata. Avvia prima install.ps1 oppure crea .venv."
}

$docsNeeded = @(
    (Join-Path $repoRoot "docs\Allegato C Creditonet (1).pdf"),
    (Join-Path $repoRoot "docs\AllegatoE determina positiva ATC (1) (3).pdf")
)
foreach ($path in $docsNeeded) {
    if (-not (Test-Path $path)) {
        throw "Template PDF mancante: $path"
    }
}

if ($Clean) {
    Write-Step "Pulizia artefatti precedenti"
    foreach ($folder in @("build", "dist")) {
        $full = Join-Path $repoRoot $folder
        if (Test-Path $full) {
            Remove-Item -LiteralPath $full -Recurse -Force
        }
    }
}

Write-Step "Verifico PyInstaller"
& $python -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Step "Installo PyInstaller nella virtualenv"
    & $python -m pip install pyinstaller
    Ensure-CommandSuccess "Installazione di PyInstaller fallita."
}

Write-Step "Genero l'eseguibile Windows"
& $python -m PyInstaller --noconfirm --clean QuintoQuote.spec
Ensure-CommandSuccess "Build PyInstaller fallita."

$distRoot = Join-Path $repoRoot "dist\QuintoQuote"
if (-not (Test-Path (Join-Path $distRoot "QuintoQuote.exe"))) {
    throw "Eseguibile non trovato dopo la build."
}

if (-not $SkipOcrBundle) {
    Write-Step "Aggiungo OCR locale al pacchetto"
    $tesseractRoot = Resolve-TesseractRoot
    Copy-TesseractBundle -SourceRoot $tesseractRoot -DestinationRoot $distRoot
}

$readmePath = Join-Path $distRoot "LEGGIMI-AVVIO.txt"
@"
QuintoQuote per Windows
=======================

1. Avvia QuintoQuote.exe con doppio click.
2. Il browser si apre automaticamente sull'app locale.
3. Config, output e immagini utente vengono salvati in:
   $env:LOCALAPPDATA\QuintoQuote
4. Per chiudere l'app usa il pulsante "Chiudi App" nella barra di navigazione.
"@ | Set-Content -LiteralPath $readmePath -Encoding UTF8

if (-not $SkipZip) {
    Write-Step "Creo archivio ZIP portabile"
    $zipPath = Join-Path $repoRoot "dist\QuintoQuote-portable.zip"
    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $distRoot "*") -DestinationPath $zipPath -CompressionLevel Optimal
}

Write-Host ""
Write-Host "Build completata." -ForegroundColor Green
Write-Host "EXE: $(Join-Path $distRoot 'QuintoQuote.exe')" -ForegroundColor Green
if (-not $SkipZip) {
    Write-Host "ZIP: $(Join-Path $repoRoot 'dist\QuintoQuote-portable.zip')" -ForegroundColor Green
}
