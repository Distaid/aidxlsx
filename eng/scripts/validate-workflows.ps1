#Requires -Version 5.1
<#
.SYNOPSIS
    Структурная проверка CI-артефактов Aid.Xlsx.

.DESCRIPTION
    Полноценного YAML-парсера на хосте нет (Windows PowerShell 5.1, pwsh отсутствует,
    PyYAML не установлен), поэтому проверки структурные:
      * наличие обязательных файлов и их базовая согласованность;
      * нет табов, отступы кратны 2;
      * uses: только из allowlist (фиксированные версии действий);
      * permissions минимальны (contents: read; id-token: write только на публикации);
      * обязательные строки и guard-ы в ci.yml (detect-projects, readiness-маркер codegen);
      * release.yml: единый путь публикации по тегу, dry-run через workflow_dispatch,
        NuGet/login@v1 + environment nuget-production, concurrency group: release,
        ancestry тега только в push-ветке, gh release create --verify-tag без --target;
      * global.json = 10.0.403 + rollForward latestPatch;
      * отсутствие подозрительных replace-вызовов в workflow и eng-скриптах;
      * отсутствие ссылок на внутренние плановые документы в коммитимых файлах
        (репозиторий публичный: ни plan-документов, ни номеров их разделов).

    Печатает PASS/FAIL по каждой проверке, затем OK либо FAIL.
    Завершается кодом 0 при успехе и 1 при любом провале.

.EXAMPLE
    powershell -NoProfile -File .\eng\scripts\validate-workflows.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\validate-workflows.ps1

.NOTES
    Код выхода: 0 — OK, 1 — есть провалы.
    На части хостов политика выполнения запрещает неподписанные .ps1: тогда используйте
    вариант с -ExecutionPolicy Bypass (это ограничение окружения, а не скрипта).
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = ''
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

$WorkflowPaths = @(
    '.github/workflows/ci.yml',
    '.github/workflows/release.yml',
    '.github/workflows/aot-matrix.yml',
    '.github/workflows/fuzz.yml'
)

$SupportPaths = @(
    'global.json',
    'Directory.Build.props',
    'Directory.Packages.props',
    '.editorconfig',
    '.gitattributes'
)

$AllowedUses = @(
    'actions/checkout@v7.0.1',
    'actions/setup-dotnet@v6.1.0',
    'actions/upload-artifact@v7.0.2',
    'actions/download-artifact@v8.0.2',
    'NuGet/login@v1'
)

$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:Passed = 0

function Add-Check {
    param(
        [string] $Name,
        [bool]   $Ok,
        [string] $Detail = ''
    )

    if ($Ok) {
        Write-Host ('PASS  ' + $Name) -ForegroundColor Green
        $script:Passed++
        return
    }

    Write-Host ('FAIL  ' + $Name) -ForegroundColor Red
    if ($Detail) {
        Write-Host ('      ' + $Detail) -ForegroundColor Yellow
        $script:Failures.Add($Name + ': ' + $Detail)
    }
    else {
        $script:Failures.Add($Name)
    }
}

function Read-RepoFile {
    param([string] $RelativePath)

    $fullPath = Join-Path $RepoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        return $null
    }
    return [System.IO.File]::ReadAllText($fullPath)
}

function Get-Lines {
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return @()
    }
    return ($Text -split "`n")
}

function Find-MissingLiteral {
    param(
        [string]   $Text,
        [string[]] $Needles
    )

    $missing = @()
    foreach ($needle in $Needles) {
        if ($Text -notmatch [regex]::Escape($needle)) {
            $missing += $needle
        }
    }
    return $missing
}

Write-Host ('Проверка CI-артефактов в ' + $RepoRoot)
Write-Host ''

# --- 1. Обязательные файлы ---------------------------------------------------
$missingFiles = @()
foreach ($relativePath in ($WorkflowPaths + $SupportPaths)) {
    if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot $relativePath))) {
        $missingFiles += $relativePath
    }
}
Add-Check 'обязательные файлы существуют' ($missingFiles.Count -eq 0) ('нет: ' + ($missingFiles -join ', '))

$workflowText = @{}
foreach ($relativePath in $WorkflowPaths) {
    $workflowText[$relativePath] = Read-RepoFile $relativePath
}

$ciText = $workflowText['.github/workflows/ci.yml']
$releaseText = $workflowText['.github/workflows/release.yml']
$aotText = $workflowText['.github/workflows/aot-matrix.yml']
$fuzzText = $workflowText['.github/workflows/fuzz.yml']

# --- 2. Нет табов ------------------------------------------------------------
$scannedPaths = $WorkflowPaths + @(
    'global.json',
    'Directory.Build.props',
    'Directory.Packages.props',
    'eng/scripts/import-upstream.ps1'
)
$tabHits = @()
foreach ($relativePath in $scannedPaths) {
    $text = Read-RepoFile $relativePath
    if ($null -eq $text) { continue }
    $lineNumber = 0
    foreach ($line in (Get-Lines $text)) {
        $lineNumber++
        if ($line.Contains([char]9)) {
            $tabHits += ($relativePath + ':' + $lineNumber)
        }
    }
}
Add-Check 'нет табов в workflow/props/скриптах' ($tabHits.Count -eq 0) ($tabHits -join ', ')

# --- 3. Отступы кратны 2 ----------------------------------------------------
$oddIndents = @()
foreach ($relativePath in $WorkflowPaths) {
    $text = $workflowText[$relativePath]
    if ($null -eq $text) { continue }
    $lineNumber = 0
    foreach ($line in (Get-Lines $text)) {
        $lineNumber++
        $trimmed = $line.TrimEnd()
        if ($trimmed.Length -eq 0) { continue }
        $leading = $trimmed.Length - $trimmed.TrimStart(' ').Length
        if (($leading -gt 0) -and (($leading % 2) -ne 0)) {
            $oddIndents += ($relativePath + ':' + $lineNumber + ' (' + $leading + ' пробелов)')
        }
    }
}
Add-Check 'отступы в YAML кратны 2' ($oddIndents.Count -eq 0) ($oddIndents -join ', ')

# --- 4. uses: только из allowlist -------------------------------------------
$badUses = @()
$usesPattern = '^\s*(?:-\s+)?uses:\s*(\S+)\s*$'
$localWorkflowPattern = '^\./\.github/workflows/[A-Za-z0-9._-]+\.ya?ml$'
foreach ($relativePath in $WorkflowPaths) {
    $text = $workflowText[$relativePath]
    if ($null -eq $text) { continue }
    $lineNumber = 0
    foreach ($line in (Get-Lines $text)) {
        $lineNumber++
        $match = [regex]::Match($line, $usesPattern)
        if (-not $match.Success) { continue }
        $use = $match.Groups[1].Value
        $allowed = $false
        foreach ($known in $AllowedUses) {
            if ($use -eq $known) { $allowed = $true }
        }
        if ((-not $allowed) -and ($use -match $localWorkflowPattern)) { $allowed = $true }
        if (-not $allowed) {
            $badUses += ($relativePath + ':' + $lineNumber + ' -> ' + $use)
        }
    }
}
Add-Check 'uses: только из allowlist' ($badUses.Count -eq 0) ($badUses -join ', ')

# --- 5. permissions ----------------------------------------------------------
$permissionIssues = @()
foreach ($relativePath in $WorkflowPaths) {
    $text = $workflowText[$relativePath]
    if ($null -eq $text) { continue }
    if ($text -notmatch '(?m)^permissions:\s*$') {
        $permissionIssues += ($relativePath + ': нет top-level permissions:')
    }
}
if ($ciText -notmatch '(?m)^permissions:\s*\r?\n\s+contents:\s*read') {
    $permissionIssues += 'ci.yml: ожидается top-level permissions с contents: read'
}
if ($releaseText -notmatch '(?m)^permissions:\s*\r?\n\s+contents:\s*read') {
    $permissionIssues += 'release.yml: ожидается top-level permissions с contents: read'
}
if ($releaseText -notmatch 'id-token:\s*write') {
    $permissionIssues += 'release.yml: нет id-token: write (Trusted Publishing)'
}
Add-Check 'permissions заданы и минимальны' ($permissionIssues.Count -eq 0) ($permissionIssues -join '; ')

# --- 6. Обязательные строки ci.yml ------------------------------------------
$ciMissing = Find-MissingLiteral -Text $ciText -Needles @(
    'name: detect-projects',
    'has_generation',
    'schemas/ooxml/SHA256SUMS',
    '*.g.cs',
    'dotnet restore',
    'dotnet format',
    '--verify-no-changes',
    '-warnaserror',
    'dotnet test',
    'Aid.Xlsx.ApiCompat',
    'Aid.Xlsx.CodeGen',
    'dotnet pack',
    'actions/upload-artifact@v7.0.2',
    'actions/setup-dotnet@v6.1.0',
    'cache: true',
    'cache-dependency-path',
    'concurrency:',
    'continue-on-error: true',
    'needs.detect.outputs.has_core',
    'needs.detect.outputs.has_api_compat',
    'needs.detect.outputs.has_aot_smoke'
)
Add-Check 'ci.yml: detect-projects и все гейты' ($ciMissing.Count -eq 0) ('нет: ' + ($ciMissing -join ', '))

# --- 7. Обязательные строки release.yml -------------------------------------
$releaseMissing = Find-MissingLiteral -Text $releaseText -Needles @(
    "tags: ['v*']",
    'workflow_dispatch:',
    'version:',
    'id-token: write',
    'NuGet/login@v1',
    'user: Distaid',
    'environment:',
    'nuget-production',
    'group: release',
    '--skip-duplicate',
    '--verify-tag',
    '.snupkg',
    'git tag --list',
    "github.event_name == 'push'",
    "github.event_name != 'push'",
    'dry-run'
)
Add-Check 'release.yml: триггеры, TP, environment, dry-run' ($releaseMissing.Count -eq 0) ('нет: ' + ($releaseMissing -join ', '))

# --- 8. global.json ----------------------------------------------------------
$globalJsonOk = $false
$globalJsonDetail = ''
$globalJsonPath = Join-Path $RepoRoot 'global.json'
if (Test-Path -LiteralPath $globalJsonPath) {
    try {
        $globalJson = [System.IO.File]::ReadAllText($globalJsonPath) | ConvertFrom-Json
        $globalJsonOk = ($globalJson.sdk.version -eq '10.0.403') -and ($globalJson.sdk.rollForward -eq 'latestPatch')
        if (-not $globalJsonOk) {
            $globalJsonDetail = ('version=' + $globalJson.sdk.version + ', rollForward=' + $globalJson.sdk.rollForward)
        }
    }
    catch {
        $globalJsonDetail = ('не читается как JSON: ' + $_.Exception.Message)
    }
}
else {
    $globalJsonDetail = 'файл отсутствует'
}
Add-Check 'global.json: SDK 10.0.403 + rollForward latestPatch' $globalJsonOk $globalJsonDetail

# --- 9. Ancestry тега только в push-ветке; dry-run через git tag --list -----
# refs/tags/$tag^{commit} допустим только внутри ветки
# github.event_name == 'push'. В dry-run тега не существует — обращение к refs/tags
# уронило бы git с exit 128, поэтому dry-run обязан использовать git tag --list.
$ancestryIssues = @()
$releaseLines = Get-Lines $releaseText
$hasAncestry = $false
for ($index = 0; $index -lt $releaseLines.Count; $index++) {
    $line = $releaseLines[$index]
    if ($line -match 'is-ancestor') { $hasAncestry = $true }
    if ($line -notmatch 'refs/tags') { continue }

    $windowStart = [Math]::Max(0, $index - 25)
    $window = ($releaseLines[$windowStart..$index] -join "`n")
    $guarded = $window -match "github\.event_name.*push|GITHUB_EVENT_NAME.*push|event_name.*'push'"
    if (-not $guarded) {
        $ancestryIssues += ('release.yml:' + ($index + 1) + ': refs/tags вне push-ветки')
    }
}
if (-not $hasAncestry) {
    $ancestryIssues += 'release.yml: нет ни одной проверки git merge-base --is-ancestor'
}
if ($releaseText -notmatch 'git tag --list') {
    $ancestryIssues += 'release.yml: dry-run не использует git tag --list'
}
Add-Check 'ancestry только в push-ветке, dry-run через git tag --list' ($ancestryIssues.Count -eq 0) ($ancestryIssues -join '; ')

# --- 10. gh release create: --verify-tag без --target -----------------------
$releaseTagIssues = @()
if ($releaseText -notmatch 'gh release create') {
    $releaseTagIssues += 'нет gh release create'
}
if ($releaseText -notmatch '--verify-tag') {
    $releaseTagIssues += 'нет --verify-tag'
}
if ($releaseText -match '--target') {
    $releaseTagIssues += 'найден --target (CI не должен управлять тегом)'
}
Add-Check 'gh release create --verify-tag, без --target' ($releaseTagIssues.Count -eq 0) ($releaseTagIssues -join '; ')

# --- 11. concurrency release ------------------------------------------------
$concurrencyOk = $false
if ($releaseText) {
    $concurrencyOk = $releaseText -match '(?m)^concurrency:\s*\r?\n\s+group:\s*release\s*\r?\n\s+cancel-in-progress:\s*false\s*$'
}
Add-Check 'release.yml: concurrency group: release, cancel-in-progress: false' $concurrencyOk ''

# --- 12. aot-matrix.yml и fuzz.yml -----------------------------------------
$aotMissing = Find-MissingLiteral -Text $aotText -Needles @(
    'cron:',
    'ubuntu-24.04',
    'windows-latest',
    'macos-latest',
    'apt-get install -y clang',
    'PublishAot=true',
    'run smoke',
    'upload-artifact@v7.0.2'
)
Add-Check 'aot-matrix.yml: матрица, clang, smoke, nightly' ($aotMissing.Count -eq 0) ('нет: ' + ($aotMissing -join ', '))

$fuzzMissing = Find-MissingLiteral -Text $fuzzText -Needles @(
    'cron:',
    'Aid.Xlsx.Fuzzing',
    'timeout 900',
    'continue-on-error: true'
)
Add-Check 'fuzz.yml: сборка targets, короткий прогон, без требования успеха' ($fuzzMissing.Count -eq 0) ('нет: ' + ($fuzzMissing -join ', '))

# --- 13. Directory.Build.props: сборка + версия + метаданные ---------------
# Версия и метаданные живут в одном Directory.Build.props (паттерн владельца, t13 rev1).
$propsText = Read-RepoFile 'Directory.Build.props'
$propsMissing = Find-MissingLiteral -Text $propsText -Needles @(
    '<LangVersion>14</LangVersion>',
    '<Nullable>enable</Nullable>',
    '<ImplicitUsings>enable</ImplicitUsings>',
    '<TreatWarningsAsErrors>true</TreatWarningsAsErrors>',
    '<Deterministic>true</Deterministic>',
    'ContinuousIntegrationBuild',
    '<EnableNETAnalyzers>true</EnableNETAnalyzers>',
    '<IsTrimmable>true</IsTrimmable>',
    '<IsAotCompatible>true</IsAotCompatible>',
    '<EnableTrimAnalyzer>true</EnableTrimAnalyzer>',
    '<EnableAotAnalyzer>true</EnableAotAnalyzer>',
    '<EnableSingleFileAnalyzer>true</EnableSingleFileAnalyzer>',
    'net10.0',
    '<VersionPrefix>0.1.0</VersionPrefix>',
    '<VersionSuffix>preview1</VersionSuffix>',
    '<Version Condition=',
    '<AssemblyVersion>',
    '<FileVersion>',
    '<Authors>Distaid</Authors>',
    '<RepositoryUrl>https://github.com/Distaid/aidxlsx</RepositoryUrl>',
    '<RepositoryType>git</RepositoryType>',
    '<PackageLicenseExpression>MIT</PackageLicenseExpression>',
    '<PublishRepositoryUrl>true</PublishRepositoryUrl>',
    '<Copyright>Copyright (c) 2026 Distaid</Copyright>'
)
Add-Check 'Directory.Build.props: сборка + версия + метаданные (TFM net10.0)' ($propsMissing.Count -eq 0) ('нет: ' + ($propsMissing -join ', '))

# Версия не должна дублироваться в отдельном файле: единственный источник — Directory.Build.props.
$versionsFileIssues = @()
if (Test-Path -LiteralPath (Join-Path $RepoRoot 'eng/Versions.props')) {
    $versionsFileIssues += 'существует eng/Versions.props — версия должна быть только в Directory.Build.props'
}
foreach ($relativePath in ($WorkflowPaths + @('Directory.Build.props', 'Directory.Packages.props'))) {
    $text = Read-RepoFile $relativePath
    if ($null -eq $text) { continue }
    if ($text -match 'Versions\.props') {
        $versionsFileIssues += ($relativePath + ': ссылается на удалённый Versions.props')
    }
}
Add-Check 'версия только в Directory.Build.props (нет Versions.props)' ($versionsFileIssues.Count -eq 0) ($versionsFileIssues -join '; ')

# release.yml обязан читать версию оттуда же, где она задана.
if ($releaseText -notmatch [regex]::Escape('Directory.Build.props')) {
    Add-Check 'release.yml читает версию из Directory.Build.props' $false 'нет ссылки на Directory.Build.props'
}
else {
    Add-Check 'release.yml читает версию из Directory.Build.props' $true
}

# --- 14. .gitattributes: бинарный корпус ------------------------------------
$gitAttributesText = Read-RepoFile '.gitattributes'
$gitAttributesIssues = @()
foreach ($extension in @('*.xlsx', '*.xlsm', '*.png', '*.bin')) {
    $pattern = '(?m)^' + [regex]::Escape($extension) + '\s+binary\b'
    if ($gitAttributesText -notmatch $pattern) {
        $gitAttributesIssues += ('нет строки ' + $extension + ' binary')
    }
}
Add-Check '.gitattributes: бинарные расширения корпуса' ($gitAttributesIssues.Count -eq 0) ($gitAttributesIssues -join '; ')

# --- 15. .editorconfig ------------------------------------------------------
$editorConfigText = Read-RepoFile '.editorconfig'
$editorConfigIssues = @()
if ($editorConfigText -notmatch '(?m)^root\s*=\s*true\s*$') { $editorConfigIssues += 'нет root = true' }
if ($editorConfigText -notmatch '(?m)^end_of_line\s*=\s*lf\s*$') { $editorConfigIssues += 'нет end_of_line = lf' }
if ($editorConfigText -notmatch '\[\*\.cs\]|\[\*\]') { $editorConfigIssues += 'нет секции для *' }
Add-Check '.editorconfig: root, единые line endings' ($editorConfigIssues.Count -eq 0) ($editorConfigIssues -join '; ')

# --- 16. Нет replace-вызовов в workflow и eng-скриптах ----------------------
$replacePattern = '(?i)re' + 'place\('
$replaceHits = @()
foreach ($relativePath in ($WorkflowPaths + @('eng/scripts/import-upstream.ps1'))) {
    $text = Read-RepoFile $relativePath
    if ($null -eq $text) { continue }
    $lineNumber = 0
    foreach ($line in (Get-Lines $text)) {
        $lineNumber++
        if ($line -match $replacePattern) {
            $replaceHits += ($relativePath + ':' + $lineNumber)
        }
    }
}
Add-Check 'нет replace-вызовов в workflow/скриптах' ($replaceHits.Count -eq 0) ($replaceHits -join ', ')

# --- 17. Trusted Publishing: основной путь без секретов ---------------------
# OIDC-публикация не требует ни API key, ни токенов: только политика TP на nuget.org
# и environment nuget-production. API key — необязательный fallback.
$secretsIssues = @()
$releaseLinesForSecrets = Get-Lines $releaseText
$trustedIndex = -1
for ($index = 0; $index -lt $releaseLinesForSecrets.Count; $index++) {
    if ($releaseLinesForSecrets[$index] -match '--api-key\s+"\$TRUSTED_API_KEY"') { $trustedIndex = $index }
}
if ($trustedIndex -lt 0) {
    $secretsIssues += 'не найден push с --api-key "$TRUSTED_API_KEY"'
}
else {
    $stepStart = 0
    for ($index = $trustedIndex; $index -ge 0; $index--) {
        if ($releaseLinesForSecrets[$index] -match '^\s*-\s+name:') { $stepStart = $index; break }
    }
    $stepEnd = $releaseLinesForSecrets.Count - 1
    for ($index = $trustedIndex; $index -lt $releaseLinesForSecrets.Count; $index++) {
        if (($index -gt $stepStart) -and ($releaseLinesForSecrets[$index] -match '^\s*-\s+name:')) { $stepEnd = $index - 1; break }
    }
    for ($index = $stepStart; $index -le $stepEnd; $index++) {
        if ($releaseLinesForSecrets[$index] -match 'secrets\.') {
            $secretsIssues += ('TP-шаг обращается к secrets (строка ' + ($index + 1) + ')')
        }
    }
}

$fallbackIndex = -1
for ($index = 0; $index -lt $releaseLinesForSecrets.Count; $index++) {
    if ($releaseLinesForSecrets[$index] -match 'secrets\.NUGET_API_KEY') { $fallbackIndex = $index }
}
if ($fallbackIndex -lt 0) {
    $secretsIssues += 'нет необязательного fallback-шага с secrets.NUGET_API_KEY'
}
else {
    $fallbackEnd = [Math]::Min($releaseLinesForSecrets.Count - 1, $fallbackIndex + 20)
    $fallbackBlock = ($releaseLinesForSecrets[$fallbackIndex..$fallbackEnd] -join "`n")
    if ($fallbackBlock -notmatch 'exit 0') {
        $secretsIssues += 'fallback не необязателен: при отсутствии секрета нет exit 0'
    }
}
Add-Check 'TP-путь без секретов; API key — необязательный fallback' ($secretsIssues.Count -eq 0) ($secretsIssues -join '; ')

# --- 18. Кодировка ----------------------------------------------------------
# Windows PowerShell 5.1 читает BOM-less не-ASCII как ANSI, поэтому props/ps1
# с кириллицей обязаны быть UTF-8 с BOM. Файлы, которые читают не PowerShell
# (.editorconfig, .gitattributes), держим чисто ASCII — BOM там вреден.
$bomIssues = @()
$bomTargets = @(
    'Directory.Build.props',
    'Directory.Packages.props',
    'eng/scripts/import-upstream.ps1',
    'eng/scripts/validate-workflows.ps1'
)
foreach ($relativePath in $bomTargets) {
    $fullPath = Join-Path $RepoRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath)) { continue }
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $hasNonAscii = $false
    foreach ($value in $bytes) {
        if ($value -gt 127) { $hasNonAscii = $true; break }
    }
    $hasBom = ($bytes.Length -ge 3) -and ($bytes[0] -eq 239) -and ($bytes[1] -eq 187) -and ($bytes[2] -eq 191)
    if ($hasNonAscii -and (-not $hasBom)) {
        $bomIssues += ($relativePath + ': не-ASCII без UTF-8 BOM (PS 5.1 прочитает как ANSI)')
    }
}
Add-Check 'кодировка: props/ps1 с не-ASCII в UTF-8 с BOM' ($bomIssues.Count -eq 0) ($bomIssues -join '; ')

$asciiIssues = @()
foreach ($relativePath in @('.editorconfig', '.gitattributes')) {
    $fullPath = Join-Path $RepoRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath)) { continue }
    foreach ($value in ([System.IO.File]::ReadAllBytes($fullPath))) {
        if ($value -gt 127) {
            $asciiIssues += ($relativePath + ': не-ASCII байт (файл должен быть ASCII)')
            break
        }
    }
}
Add-Check '.editorconfig/.gitattributes — ASCII без BOM' ($asciiIssues.Count -eq 0) ($asciiIssues -join '; ')

# --- 19. Нет ссылок на плановые документы в коммитимых файлах --------------
# Репозиторий публичный: коммитимые файлы не должны ссылаться на внутренние
# плановые документы, research-файлы и номера их разделов.
# Маркеры собираются из фрагментов (а знак раздела — из кода символа), чтобы
# сам этот скрипт не находил маркеры в себе и не ловил ложные срабатывания.
#
# Ссылки вида §4.1 / §12/O16 (наша нумерация разделов плана) ловятся регулярками;
# цитаты спецификаций вида §18.8.30 (ECMA-376) намеренно НЕ ловятся — это
# легитимные ссылки в исходниках upstream, их нельзя переписывать.
$sectionSign = [string][char]0x00A7
$planningMarkers = @(
    ('docs' + '/plan'),
    ('IMPLEMENTATION' + '-PLAN'),
    ('PLAN' + '.md'),
    ('docs' + '/research'),
    ('[' + 'facts'),
    ('плана' + ' ' + $sectionSign)
)
$planningPatterns = @(
    $sectionSign + '[1-9]\.',
    $sectionSign + '\d{1,2}/[A-Z]'
)

$committableFiles = @()
if ($null -ne (Get-Command git -ErrorAction SilentlyContinue)) {
    $tracked = @(& git -C $RepoRoot ls-files 2>$null)
    $untracked = @(& git -C $RepoRoot ls-files --others --exclude-standard 2>$null)
    $committableFiles = @(($tracked + $untracked) | Sort-Object -Unique)
}
if ($committableFiles.Count -eq 0) {
    # Fallback без git: известный набор файлов, которые попадут в репозиторий.
    $committableFiles = @(
        'global.json',
        'Directory.Build.props',
        'Directory.Packages.props',
        '.editorconfig',
        '.gitattributes',
        '.gitignore',
        'README.md',
        'CHANGELOG.md'
    ) + $WorkflowPaths + @(
        'eng/scripts/import-upstream.ps1',
        'eng/scripts/validate-workflows.ps1'
    )
}

$textExtensions = @(
    '.md', '.txt', '.yml', '.yaml', '.json', '.props', '.targets', '.csproj',
    '.sln', '.slnx', '.ps1', '.psm1', '.cs', '.config', '.sh', '.cmd', '.bat', '.xml'
)
$extensionlessTextFiles = @('.editorconfig', '.gitattributes', '.gitignore')

$planningHits = @()
foreach ($relativePath in $committableFiles) {
    $fullPath = Join-Path $RepoRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath)) { continue }

    $fileName = [System.IO.Path]::GetFileName($relativePath).ToLowerInvariant()
    $extension = [System.IO.Path]::GetExtension($relativePath).ToLowerInvariant()
    $isText = ($textExtensions -contains $extension) -or ($extensionlessTextFiles -contains $fileName)
    if (-not $isText) { continue }

    $lineNumber = 0
    foreach ($line in [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $fullPath))) {
        $lineNumber++
        $matched = $false
        foreach ($marker in $planningMarkers) {
            if ($line.IndexOf($marker, [System.StringComparison]::Ordinal) -ge 0) {
                $matched = $true
                break
            }
        }
        if (-not $matched) {
            foreach ($pattern in $planningPatterns) {
                if ($line -match $pattern) {
                    $matched = $true
                    break
                }
            }
        }
        if ($matched) {
            $planningHits += ('ссылка на плановый документ в публичном файле: ' + $relativePath + ':' + $lineNumber)
        }
    }
}
# НЕ ПРОВЕРЯЮТСЯ: .git/info/exclude и локальные плановые документы — они не коммитятся
# (исключены через .git/info/exclude), поэтому git ls-files их не возвращает.
Add-Check 'нет ссылок на плановые документы в коммитимых файлах' ($planningHits.Count -eq 0) ($planningHits -join '; ')

# --- Итог -------------------------------------------------------------------
Write-Host ''
if ($script:Failures.Count -eq 0) {
    Write-Host ('OK: ' + $script:Passed + ' проверок пройдено') -ForegroundColor Green
    exit 0
}

Write-Host ('FAIL: провалено ' + $script:Failures.Count + ' из ' + ($script:Passed + $script:Failures.Count) + ' проверок') -ForegroundColor Red
foreach ($failure in $script:Failures) {
    Write-Host ('  - ' + $failure) -ForegroundColor Red
}
exit 1