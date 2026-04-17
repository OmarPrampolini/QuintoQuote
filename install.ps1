param(
    [string]$TargetDir = (Join-Path $HOME "QuintoQuote"),
    [string]$Branch = "main",
    [string]$RepoUrl = "https://github.com/OmarPrampolini/QuintoQuote.git",
    [string]$ArchiveUrl = "",
    [string]$SourcePath = "",
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
    Write-Host "==> $Message" -ForegroundColor Cyan
}

if (
    -not $SourcePath `
    -and -not $PSBoundParameters.ContainsKey("TargetDir") `
    -and $PSScriptRoot `
    -and (Test-Path (Join-Path $PSScriptRoot "pyproject.toml"))
) {
    $SourcePath = $PSScriptRoot
    if ($TargetDir -eq (Join-Path $HOME "QuintoQuote")) {
        $TargetDir = $PSScriptRoot
    }
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

function Resolve-ArchiveUrl([string]$GitUrl, [string]$BranchName, [string]$ExplicitUrl) {
    if ($ExplicitUrl) {
        return $ExplicitUrl
    }
    $baseUrl = $GitUrl
    if ($baseUrl.EndsWith(".git")) {
        $baseUrl = $baseUrl.Substring(0, $baseUrl.Length - 4)
    }
    return "$baseUrl/archive/refs/heads/$BranchName.zip"
}

function Write-AsciiFile([string]$Path, [string[]]$Lines) {
    $content = ($Lines -join "`r`n") + "`r`n"
    $changed = $true
    if (Test-Path $Path) {
        $existing = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::ASCII)
        $changed = $existing -ne $content
    }
    if ($changed) {
        [System.IO.File]::WriteAllText($Path, $content, [System.Text.Encoding]::ASCII)
    }
    return $changed
}

function Get-UserBinDir {
    return (Join-Path $env:LOCALAPPDATA "QuintoQuote\bin")
}

function Get-LocalTesseractRoot {
    return (Join-Path $env:LOCALAPPDATA "QuintoQuote\ocr\Tesseract-OCR")
}

function Get-LocalTesseractExecutable {
    return (Join-Path (Get-LocalTesseractRoot) "tesseract.exe")
}

function Set-UserEnvironmentVariable([string]$Name, [string]$Value) {
    [Environment]::SetEnvironmentVariable($Name, $Value, "User")
    Set-Item -Path ("Env:" + $Name) -Value $Value
}

function Ensure-UserPathContains([string]$Entry) {
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $parts = @()
    if ($userPath) {
        $parts = @($userPath -split ";" | Where-Object { $_ })
    }
    if ($parts -contains $Entry) {
        if (-not (($env:Path -split ";") -contains $Entry)) {
            $env:Path = "$Entry;$env:Path"
        }
        return $false
    }
    $newUserPath = if ($userPath) { "$userPath;$Entry" } else { $Entry }
    [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
    if (-not (($env:Path -split ";") -contains $Entry)) {
        $env:Path = "$Entry;$env:Path"
    }
    return $true
}

function Install-UserCommands([string]$RepoRoot, [string]$TesseractExecutablePath = "") {
    $binDir = Get-UserBinDir
    New-Item -ItemType Directory -Force -Path $binDir | Out-Null

    $launcherCmd = Join-Path $binDir "quintoquote.cmd"
    $updateCmd = Join-Path $binDir "quintoquote-update.cmd"
    $repoRootWindows = $RepoRoot
    $installScript = Join-Path $RepoRoot "install.ps1"
    $tessdataDir = ""
    if ($TesseractExecutablePath) {
        $tessdataDir = Get-TesseractTessdataDir -ExecutablePath $TesseractExecutablePath
    }

    $launcherChanged = Write-AsciiFile -Path $launcherCmd -Lines @(
        "@echo off",
        "setlocal",
        "set ""QUINTOQUOTE_ROOT=$repoRootWindows""",
        "set ""QUINTOQUOTE_LAUNCHER=%QUINTOQUOTE_ROOT%\.venv\Scripts\quintoquote.exe""",
        $(if ($TesseractExecutablePath) { "set ""QUINTOQUOTE_TESSERACT_PATH=$TesseractExecutablePath""" }),
        $(if ($tessdataDir) { "set ""TESSDATA_PREFIX=$tessdataDir""" }),
        "if not exist ""%QUINTOQUOTE_LAUNCHER%"" (",
        "  echo QuintoQuote non trovato in %QUINTOQUOTE_ROOT%. Esegui di nuovo l'installer.",
        "  exit /b 1",
        ")",
        "if ""%~1""=="""" goto runstart",
        "set ""QQ_FIRST_ARG=%~1""",
        "if ""%QQ_FIRST_ARG:~0,1%""==""-"" goto runstart",
        """%QUINTOQUOTE_LAUNCHER%"" %*",
        "exit /b %ERRORLEVEL%",
        ":runstart",
        """%QUINTOQUOTE_LAUNCHER%"" start %*"
    )

    $updateChanged = Write-AsciiFile -Path $updateCmd -Lines @(
        "@echo off",
        "setlocal",
        $(if ($TesseractExecutablePath) { "set ""QUINTOQUOTE_TESSERACT_PATH=$TesseractExecutablePath""" }),
        $(if ($tessdataDir) { "set ""TESSDATA_PREFIX=$tessdataDir""" }),
        "powershell -ExecutionPolicy Bypass -File ""$installScript"" -TargetDir ""$repoRootWindows"" -NoLaunch %*"
    )

    $pathAdded = Ensure-UserPathContains -Entry $binDir
    return [pscustomobject]@{
        BinDir = $binDir
        PathAdded = $pathAdded
        LauncherCmd = $launcherCmd
        UpdateCmd = $updateCmd
        LauncherChanged = [bool]$launcherChanged
        UpdateChanged = [bool]$updateChanged
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

function Invoke-GitCapture([string]$GitExe, [string[]]$Arguments) {
    $output = & $GitExe @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = [string]::Join("`n", @($output))
    }
}

function Sync-RepoFromArchive([string]$Destination, [string]$ZipUrl) {
    Write-Step "Sincronizzo la cartella dal pacchetto remoto"
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
    $preserve = @(".venv", ".git", ".quintoquote_tmp", "output_preventivi", "assets", "config.json")
    $sourceNames = @{}
    Get-ChildItem -LiteralPath $expanded.FullName -Force | ForEach-Object {
        $sourceNames[$_.Name] = $true
    }
    Get-ChildItem -LiteralPath $Destination -Force | ForEach-Object {
        if ($preserve -contains $_.Name) {
            return
        }
        if (-not $sourceNames.ContainsKey($_.Name)) {
            Remove-Item -LiteralPath $_.FullName -Recurse -Force
        }
    }
    Get-ChildItem -LiteralPath $expanded.FullName -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
    }
}

function Ensure-Repo([string]$Destination, [string]$BranchName, [string]$GitUrl, [string]$ZipUrl, [string]$LocalSource) {
    if ($LocalSource) {
        $resolvedSource = (Resolve-Path $LocalSource).Path
        $resolvedDestination = [System.IO.Path]::GetFullPath($Destination)
        if (Test-Path $Destination) {
            $resolvedDestination = (Resolve-Path $Destination).Path
        }
        if ($resolvedSource -eq $resolvedDestination) {
            Write-Step "Uso la cartella locale esistente in $resolvedSource"
            return [pscustomobject]@{
                Path = $resolvedSource
                Changed = $false
                Source = "local"
                Detail = "Checkout locale gia in uso."
            }
        }
        Write-Step "Copio il repository locale in $Destination"
        Sync-LocalSource -FromPath $LocalSource -ToPath $Destination
        return [pscustomobject]@{
            Path = [System.IO.Path]::GetFullPath($Destination)
            Changed = $true
            Source = "local-copy"
            Detail = "Repository copiato da sorgente locale."
        }
    }

    if (Test-Path (Join-Path $Destination ".git")) {
        $git = Get-Command git -ErrorAction SilentlyContinue
        if ($git) {
            Write-Step "Controllo allineamento con la repo remota"
            $fetch = Invoke-GitCapture $git.Source @("-C", $Destination, "fetch", "--prune", "origin", $BranchName)
            if ($fetch.ExitCode -ne 0) {
                throw "Fetch git fallito: $($fetch.Output)"
            }
            $head = (Invoke-GitCapture $git.Source @("-C", $Destination, "rev-parse", "HEAD")).Output.Trim()
            $remote = (Invoke-GitCapture $git.Source @("-C", $Destination, "rev-parse", "origin/$BranchName")).Output.Trim()
            if (-not $head -or -not $remote) {
                throw "Impossibile determinare lo stato del branch remoto."
            }
            if ($head -eq $remote) {
                return [pscustomobject]@{
                    Path = (Resolve-Path $Destination).Path
                    Changed = $false
                    Source = "git"
                    Detail = "Repository gia allineato a origin/$BranchName."
                }
            }

            $dirty = (Invoke-GitCapture $git.Source @("-C", $Destination, "status", "--porcelain")).Output.Trim()
            if ($dirty) {
                throw "La cartella contiene modifiche locali non salvate. Salvale o ripristinale prima di eseguire quintoquote-update."
            }

            Write-Step "Aggiorno la cartella alla versione remota"
            $reset = Invoke-GitCapture $git.Source @("-C", $Destination, "reset", "--hard", "origin/$BranchName")
            if ($reset.ExitCode -ne 0) {
                throw "Reset git fallito: $($reset.Output)"
            }
            return [pscustomobject]@{
                Path = (Resolve-Path $Destination).Path
                Changed = $true
                Source = "git"
                Detail = "Repository riallineato a origin/$BranchName."
            }
        }
    }

    if (Test-Path (Join-Path $Destination "pyproject.toml")) {
        Sync-RepoFromArchive -Destination $Destination -ZipUrl $ZipUrl
        return [pscustomobject]@{
            Path = (Resolve-Path $Destination).Path
            Changed = $true
            Source = "archive-refresh"
            Detail = "Cartella sincronizzata dall'archivio remoto."
        }
    }

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git) {
        Write-Step "Clono il repository GitHub"
        & $git.Source clone --depth 1 --branch $BranchName $GitUrl $Destination
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{
                Path = (Resolve-Path $Destination).Path
                Changed = $true
                Source = "git-clone"
                Detail = "Repository clonato da GitHub."
            }
        }
    }

    Sync-RepoFromArchive -Destination $Destination -ZipUrl $ZipUrl
    return [pscustomobject]@{
        Path = (Resolve-Path $Destination).Path
        Changed = $true
        Source = "archive"
        Detail = "Repository installato dall'archivio remoto."
    }
}

function Get-TesseractRootFromExecutable([string]$ExecutablePath) {
    if (-not $ExecutablePath) {
        return ""
    }
    $resolved = (Resolve-Path $ExecutablePath -ErrorAction SilentlyContinue)
    if (-not $resolved) {
        return ""
    }
    return Split-Path -Parent $resolved.Path
}

function Find-TesseractExecutable {
    $envConfigured = [Environment]::GetEnvironmentVariable("QUINTOQUOTE_TESSERACT_PATH", "User")
    $candidates = @(
        $envConfigured,
        (Get-LocalTesseractExecutable),
        "C:\Program Files\Tesseract-OCR\tesseract.exe",
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($tesseractCmd) {
        return $tesseractCmd.Source
    }
    return ""
}

function Get-TesseractTessdataDir([string]$ExecutablePath) {
    $root = Get-TesseractRootFromExecutable -ExecutablePath $ExecutablePath
    if (-not $root) {
        return ""
    }
    $candidates = @(
        (Join-Path $root "tessdata"),
        (Join-Path (Split-Path -Parent $root) "tessdata")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }
    return (Join-Path $root "tessdata")
}

function Assert-TrustedModelSource([string]$Url, [string]$Name) {
    $uri = [System.Uri]$Url
    if ($uri.Scheme -ne "https") {
        throw "Origine non valida per ${Name}: e richiesto HTTPS."
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

function Ensure-TesseractLanguageData([string]$ExecutablePath) {
    $downloaded = 0
    try {
        $tessdataDir = Get-TesseractTessdataDir -ExecutablePath $ExecutablePath
        if (-not $tessdataDir) {
            throw "Cartella tessdata non trovata per Tesseract."
        }
        New-Item -ItemType Directory -Force -Path $tessdataDir | Out-Null

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
            $targetFile = Join-Path $tessdataDir $model.Name
            if (Test-Path $targetFile) {
                continue
            }
            Write-Step "Scarico modello OCR $($model.Name)"
            Install-TessdataModel -Url $model.Url -TargetFile $targetFile -ExpectedSha256 $model.Sha256
            $downloaded += 1
        }
    } catch {
        Write-Host "Avviso: non sono riuscito a completare i modelli OCR ($($_.Exception.Message))." -ForegroundColor Yellow
    }
    return [pscustomobject]@{
        ExecutablePath = $ExecutablePath
        DownloadedModels = $downloaded
        Changed = ($downloaded -gt 0)
    }
}

function Get-LatestTesseractInstallerUrl {
    $indexUrl = "https://digi.bib.uni-mannheim.de/tesseract/"
    $response = Invoke-WebRequest -Uri $indexUrl -UseBasicParsing
    $matches = [regex]::Matches($response.Content, 'href="(tesseract-ocr-w64-setup-[^"]+\.exe)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $candidates = @()
    foreach ($match in $matches) {
        $name = $match.Groups[1].Value
        if (-not $name) {
            continue
        }
        if ($name -match '(?i)(alpha|beta|rc|dev)') {
            continue
        }
        $dateKey = "00000000"
        if ($name -match '(\d{8})(?=\.exe$)') {
            $dateKey = $Matches[1]
        }
        $candidates += [pscustomobject]@{
            Name = $name
            DateKey = $dateKey
            Url = "$indexUrl$name"
        }
    }
    if (-not $candidates) {
        throw "Nessun installer Tesseract stabile trovato nella directory ufficiale UB Mannheim."
    }
    $selected = $candidates | Sort-Object DateKey, Name -Descending | Select-Object -First 1
    return $selected.Url
}

function Install-TesseractFromOfficialInstaller {
    $installerUrl = Get-LatestTesseractInstallerUrl
    Write-Step "Scarico Tesseract OCR dal catalogo ufficiale UB Mannheim"
    $installerPath = Join-Path $env:TEMP "qq_tesseract_installer.exe"
    $localRoot = Get-LocalTesseractRoot
    New-Item -ItemType Directory -Force -Path $localRoot | Out-Null
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
    try {
        $arguments = @(
            "/SP-",
            "/VERYSILENT",
            "/SUPPRESSMSGBOXES",
            "/NORESTART",
            "/CURRENTUSER",
            "/DIR=$localRoot"
        )
        $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "Installer Tesseract terminato con exit code $($process.ExitCode) su $localRoot."
        }
    } finally {
        if (Test-Path $installerPath) {
            Remove-Item -LiteralPath $installerPath -Force
        }
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

    & $winget.Source install --id UB-Mannheim.TesseractOCR --exact --silent --accept-package-agreements --accept-source-agreements
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

function Ensure-TesseractRobust {
    Write-Step "Controllo Tesseract OCR"
    $existingExecutable = Find-TesseractExecutable
    if ($existingExecutable) {
        Write-Host "Tesseract OCR gia installato." -ForegroundColor Green
        $langInfo = Ensure-TesseractLanguageData -ExecutablePath $existingExecutable
        Set-UserEnvironmentVariable -Name "QUINTOQUOTE_TESSERACT_PATH" -Value $existingExecutable
        $tessdataDir = Get-TesseractTessdataDir -ExecutablePath $existingExecutable
        if ($tessdataDir) {
            Set-UserEnvironmentVariable -Name "TESSDATA_PREFIX" -Value $tessdataDir
        }
        return [pscustomobject]@{
            ExecutablePath = $existingExecutable
            Source = "existing"
            Changed = [bool]$langInfo.Changed
            DownloadedModels = [int]$langInfo.DownloadedModels
        }
    }

    $installedBy = "winget"
    Write-Step "Tesseract OCR non trovato. Tento l'installazione automatica..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        & $winget.Source install --id UB-Mannheim.TesseractOCR --exact --silent --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Installazione via winget non riuscita, provo con l'installer ufficiale." -ForegroundColor Yellow
            $installedBy = "official-installer"
        }
    } else {
        $installedBy = "official-installer"
    }

    $installedExecutable = Find-TesseractExecutable
    if (-not $installedExecutable) {
        try {
            Install-TesseractFromOfficialInstaller
            $installedBy = "official-installer"
        } catch {
            $message = $_.Exception.Message
            throw "Installazione automatica di Tesseract OCR fallita: $message"
        }
    }

    $installedExecutable = Find-TesseractExecutable
    if ($installedExecutable) {
        $langInfo = Ensure-TesseractLanguageData -ExecutablePath $installedExecutable
        Set-UserEnvironmentVariable -Name "QUINTOQUOTE_TESSERACT_PATH" -Value $installedExecutable
        $tessdataDir = Get-TesseractTessdataDir -ExecutablePath $installedExecutable
        if ($tessdataDir) {
            Set-UserEnvironmentVariable -Name "TESSDATA_PREFIX" -Value $tessdataDir
        }
        return [pscustomobject]@{
            ExecutablePath = $installedExecutable
            Source = $installedBy
            Changed = $true
            DownloadedModels = [int]$langInfo.DownloadedModels
        }
    }

    throw "Tesseract OCR installato ma non trovato nel PATH. Riavvia la shell o aggiungi manualmente il percorso di installazione al PATH."
}

function Write-DoctorSummary(
    $RepoInfo,
    $TesseractInfo,
    [bool]$VenvCreated,
    [bool]$LauncherRecreated,
    $CommandInfo,
    [bool]$PipHealthy,
    [bool]$ImportHealthy
) {
    $changes = @()
    if ($RepoInfo.Changed) {
        $changes += $RepoInfo.Detail
    }
    if ($TesseractInfo.Changed) {
        if ($TesseractInfo.Source -eq "existing") {
            $changes += "OCR verificato e completato con $($TesseractInfo.DownloadedModels) modelli mancanti."
        } else {
            $changes += "OCR reinstallato o configurato automaticamente ($($TesseractInfo.Source))."
        }
    }
    if ($VenvCreated) {
        $changes += "Virtualenv ricreata o inizializzata."
    }
    if ($LauncherRecreated) {
        $changes += "Launcher locale ricreato."
    }
    if ($CommandInfo.LauncherChanged -or $CommandInfo.UpdateChanged) {
        $changes += "Comandi globali aggiornati."
    }
    if ($CommandInfo.PathAdded) {
        $changes += "PATH utente aggiornato."
    }

    Write-Host ""
    if ($changes.Count -eq 0 -and $PipHealthy -and $ImportHealthy) {
        Write-Host "Tutto in linea: QuintoQuote e gia allineato, non devi fare nulla." -ForegroundColor Green
    } else {
        Write-Host "QuintoQuote aggiornato e verificato." -ForegroundColor Green
        foreach ($change in $changes) {
            Write-Host " - $change" -ForegroundColor DarkGray
        }
    }

    if (-not $PipHealthy) {
        Write-Host "Avviso: pip check non e andato a buon fine." -ForegroundColor Yellow
    }
    if (-not $ImportHealthy) {
        Write-Host "Avviso: il test di import del package non e andato a buon fine." -ForegroundColor Yellow
    }
}

$python = Resolve-PythonCommand
$resolvedArchiveUrl = Resolve-ArchiveUrl -GitUrl $RepoUrl -BranchName $Branch -ExplicitUrl $ArchiveUrl
$tesseractInfo = Ensure-TesseractRobust
$repoInfo = Ensure-Repo -Destination $TargetDir -BranchName $Branch -GitUrl $RepoUrl -ZipUrl $resolvedArchiveUrl -LocalSource $SourcePath

$repoRoot = (Resolve-Path $TargetDir).Path
$runtimeTmp = Join-Path $repoRoot ".quintoquote_tmp"
New-Item -ItemType Directory -Force -Path $runtimeTmp | Out-Null
$env:TEMP = $runtimeTmp
$env:TMP = $runtimeTmp

Write-Step "Creo o aggiorno la virtualenv"
$venvPath = Join-Path $repoRoot ".venv"
$venvPython = Join-Path $venvPath "Scripts\\python.exe"
$venvExistedBefore = Test-Path $venvPython
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

$launcher = Join-Path $repoRoot ".venv\\Scripts\\quintoquote.exe"
$launcherExistedBefore = Test-Path $launcher

Write-Step "Installo QuintoQuote"
Push-Location $repoRoot
try {
    Invoke-Python $python @("-m", "pip", "--python", $venvPath, "install", "-e", ".")
} finally {
    Pop-Location
}

if (-not (Test-Path $launcher)) {
    throw "Launcher non trovato: $launcher"
}

Write-Step "Verifico dipendenze e package"
$pipHealthy = $true
try {
    Invoke-Python $python @("-m", "pip", "--python", $venvPath, "check")
} catch {
    $pipHealthy = $false
}

$importHealthy = $true
try {
    & $venvPython -c "import preventivo_generator_v2"
    if ($LASTEXITCODE -ne 0) {
        throw "Import package failed"
    }
} catch {
    $importHealthy = $false
}

$commandInfo = Install-UserCommands -RepoRoot $repoRoot -TesseractExecutablePath $tesseractInfo.ExecutablePath
$venvCreated = -not $venvExistedBefore
$launcherRecreated = (-not $launcherExistedBefore) -and (Test-Path $launcher)

Write-Host ""
Write-Host "QuintoQuote installato in $repoRoot" -ForegroundColor Green
Write-Host "OCR locale attivo: $($tesseractInfo.ExecutablePath)" -ForegroundColor Green
Write-Host "Comando avvio: quintoquote" -ForegroundColor DarkGray
Write-Host "Comando update: quintoquote-update" -ForegroundColor DarkGray
if ($commandInfo.PathAdded) {
    Write-Host "PATH utente aggiornato: riapri il terminale per usare i comandi globali." -ForegroundColor Yellow
}
Write-DoctorSummary -RepoInfo $repoInfo -TesseractInfo $tesseractInfo -VenvCreated $venvCreated -LauncherRecreated $launcherRecreated -CommandInfo $commandInfo -PipHealthy $pipHealthy -ImportHealthy $importHealthy
Write-Host ""

if (-not $NoLaunch) {
    Write-Step "Avvio QuintoQuote"
    & $launcher start
}
