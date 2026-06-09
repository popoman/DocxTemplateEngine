$ErrorActionPreference = 'Stop'

$invokeDir   = $PSScriptRoot
$outputFile  = 'res.docx'
$outputPath  = Join-Path $invokeDir $outputFile
$templateDoc = Join-Path $invokeDir 'template.docx'
$configFile  = Join-Path $invokeDir 'config.yaml'

# Resolve the freshest local Debug build of DocxTemplateEngine.exe in ../src.
# Prefer win-x64\ since the .csproj sets <RuntimeIdentifier>win-x64</RuntimeIdentifier>;
# fall back to net8.0\ for plain `dotnet build` without RID.
$repoRoot = Split-Path $invokeDir -Parent
$candidates = @(
    Join-Path $repoRoot 'src\DocxTemplateEngine\bin\Debug\net8.0\win-x64\DocxTemplateEngine.exe'
    Join-Path $repoRoot 'src\DocxTemplateEngine\bin\Debug\net8.0\DocxTemplateEngine.exe'
)
$exe = $candidates |
    Where-Object { Test-Path $_ } |
    Sort-Object { (Get-Item $_).LastWriteTime } -Descending |
    Select-Object -First 1
if (-not $exe) {
    throw "No Debug build of DocxTemplateEngine.exe found under $repoRoot\src. Run: dotnet build src/DocxTemplateEngine"
}
Write-Host "Using exe: $exe"

# Close the output document in Word if it's already open. Try CommandLine first
# (precise but Win32_Process.CommandLine is empty for processes whose security
# context we can't read), fall back to MainWindowTitle (catches e.g. "res.docx - Word").
$pidsToKill = New-Object System.Collections.Generic.HashSet[int]
Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
    Where-Object { $_.CommandLine -like "*$outputFile*" } |
    ForEach-Object { [void]$pidsToKill.Add([int]$_.ProcessId) }
Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$outputFile*" } |
    ForEach-Object { [void]$pidsToKill.Add([int]$_.Id) }
if ($pidsToKill.Count -gt 0) {
    Write-Host "Closing $outputFile in Word (PIDs: $($pidsToKill -join ', '))..."
    foreach ($wordPid in $pidsToKill) { Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue }
    # Word releases the file lock asynchronously; wait briefly.
    Start-Sleep -Milliseconds 300
}

# Delete any previous output so we never open a stale file on engine failure.
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}

# Generate. Push into invoke/ so config.yaml's relative source paths resolve.
Push-Location $invokeDir
try {
    Write-Host "Building $outputFile..."
    & $exe -t $templateDoc -c $configFile -o $outputFile
    if ($LASTEXITCODE -ne 0) {
        throw "DocxTemplateEngine exited with code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

Start-Process $outputPath
