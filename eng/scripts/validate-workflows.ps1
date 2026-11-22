#Requires -Version 5.1
<#
.SYNOPSIS
    Structural checks for the Aid.Xlsx CI artifacts.

.DESCRIPTION
    There is no full YAML parser on the host (Windows PowerShell 5.1, no pwsh,
    PyYAML is not installed), so the checks are structural:
      * required files exist and are basically consistent;
      * no tabs, indentation is a multiple of 2;
      * uses: only from the allowlist (pinned action versions);
      * permissions are minimal (contents: read; id-token: write only for publishing);
      * required lines and guards in ci.yml (detect-projects, codegen readiness marker);
      * release.yml: single publication path by tag, dry-run via workflow_dispatch,
        NuGet/login@v1 + environment nuget-production, concurrency group: release,
        tag ancestry only in the push branch, gh release create --verify-tag without --target;
      * global.json pins SDK 10.0.403 with rollForward disabled (an exact pin keeps lock files stable);
      * no suspicious replace calls in workflows and eng scripts;
      * no references to internal planning documents in committable files
        (the repository is public: no plan documents, no section numbers).

    Prints PASS/FAIL per check, then OK or FAIL.
    Exits with code 0 on success and 1 on any failure.

.EXAMPLE
    powershell -NoProfile -File .\eng\scripts\validate-workflows.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\validate-workflows.ps1

.NOTES
    Exit code: 0 - OK, 1 - there are failures.
    On some hosts the execution policy blocks unsigned .ps1 files: use the
    -ExecutionPolicy Bypass variant (an environment limitation, not a script problem).
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

Write-Host ('Checking CI artifacts in ' + $RepoRoot)
Write-Host ''

# --- 1. Required files --------------------------------------------------
$missingFiles = @()
foreach ($relativePath in ($WorkflowPaths + $SupportPaths)) {
    if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot $relativePath))) {
        $missingFiles += $relativePath
    }
}
Add-Check 'required files exist' ($missingFiles.Count -eq 0) ('missing: ' + ($missingFiles -join ', '))

$workflowText = @{}
foreach ($relativePath in $WorkflowPaths) {
    $workflowText[$relativePath] = Read-RepoFile $relativePath
}

$ciText = $workflowText['.github/workflows/ci.yml']
$releaseText = $workflowText['.github/workflows/release.yml']
$aotText = $workflowText['.github/workflows/aot-matrix.yml']
$fuzzText = $workflowText['.github/workflows/fuzz.yml']

# --- 2. No tabs --------------------------------------------------------
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
Add-Check 'no tabs in workflows/props/scripts' ($tabHits.Count -eq 0) ($tabHits -join ', ')

# --- 3. Indentation is a multiple of 2 ---------------------------------
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
            $oddIndents += ($relativePath + ':' + $lineNumber + ' (' + $leading + ' spaces)')
        }
    }
}
Add-Check 'YAML indentation is a multiple of 2' ($oddIndents.Count -eq 0) ($oddIndents -join ', ')

# --- 4. uses: allowlist only -------------------------------------------
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
Add-Check 'uses: allowlist only' ($badUses.Count -eq 0) ($badUses -join ', ')

# --- 5. permissions ----------------------------------------------------------
$permissionIssues = @()
foreach ($relativePath in $WorkflowPaths) {
    $text = $workflowText[$relativePath]
    if ($null -eq $text) { continue }
    if ($text -notmatch '(?m)^permissions:\s*$') {
        $permissionIssues += ($relativePath + ': no top-level permissions:')
    }
}
if ($ciText -notmatch '(?m)^permissions:\s*\r?\n\s+contents:\s*read') {
    $permissionIssues += 'ci.yml: expected top-level permissions with contents: read'
}
if ($releaseText -notmatch '(?m)^permissions:\s*\r?\n\s+contents:\s*read') {
    $permissionIssues += 'release.yml: expected top-level permissions with contents: read'
}
if ($releaseText -notmatch 'id-token:\s*write') {
    $permissionIssues += 'release.yml: no id-token: write (Trusted Publishing)'
}
Add-Check 'permissions are set and minimal' ($permissionIssues.Count -eq 0) ($permissionIssues -join '; ')

# --- 6. Required lines in ci.yml ---------------------------------------
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
Add-Check 'ci.yml: detect-projects and all gates' ($ciMissing.Count -eq 0) ('missing: ' + ($ciMissing -join ', '))

# --- 7. Required lines in release.yml ----------------------------------
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
Add-Check 'release.yml: triggers, TP, environment, dry-run' ($releaseMissing.Count -eq 0) ('missing: ' + ($releaseMissing -join ', '))

# --- 8. global.json ----------------------------------------------------------
$globalJsonOk = $false
$globalJsonDetail = ''
$globalJsonPath = Join-Path $RepoRoot 'global.json'
if (Test-Path -LiteralPath $globalJsonPath) {
    try {
        $globalJson = [System.IO.File]::ReadAllText($globalJsonPath) | ConvertFrom-Json
        # The SDK is pinned exactly (rollForward disabled): a floating SDK would change restore
        # behaviour and the lock files between machines.
        $globalJsonOk = ($globalJson.sdk.version -eq '10.0.403') -and ($globalJson.sdk.rollForward -eq 'disable')
        if (-not $globalJsonOk) {
            $globalJsonDetail = ('version=' + $globalJson.sdk.version + ', rollForward=' + $globalJson.sdk.rollForward)
        }
    }
    catch {
        $globalJsonDetail = ('not readable as JSON: ' + $_.Exception.Message)
    }
}
else {
    $globalJsonDetail = 'file is missing'
}
Add-Check 'global.json: SDK 10.0.403 with rollForward disabled' $globalJsonOk $globalJsonDetail

# --- 9. Tag ancestry only in the push branch; dry-run via git tag --list
# refs/tags/$tag^{commit} is allowed only inside the branch
# github.event_name == 'push'. In a dry run the tag does not exist, so touching refs/tags
# would make git exit 128; a dry run must use git tag --list instead.
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
        $ancestryIssues += ('release.yml:' + ($index + 1) + ': refs/tags outside the push branch')
    }
}
if (-not $hasAncestry) {
    $ancestryIssues += 'release.yml: no git merge-base --is-ancestor check at all'
}
if ($releaseText -notmatch 'git tag --list') {
    $ancestryIssues += 'release.yml: dry-run does not use git tag --list'
}
Add-Check 'ancestry only in the push branch, dry-run via git tag --list' ($ancestryIssues.Count -eq 0) ($ancestryIssues -join '; ')

# --- 10. gh release create: --verify-tag without --target ---------------
$releaseTagIssues = @()
if ($releaseText -notmatch 'gh release create') {
    $releaseTagIssues += 'no gh release create'
}
if ($releaseText -notmatch '--verify-tag') {
    $releaseTagIssues += 'no --verify-tag'
}
if ($releaseText -match '--target') {
    $releaseTagIssues += '--target found (CI must not manage the tag)'
}
Add-Check 'gh release create --verify-tag, no --target' ($releaseTagIssues.Count -eq 0) ($releaseTagIssues -join '; ')

# --- 11. concurrency release ------------------------------------------------
$concurrencyOk = $false
if ($releaseText) {
    $concurrencyOk = $releaseText -match '(?m)^concurrency:\s*\r?\n\s+group:\s*release\s*\r?\n\s+cancel-in-progress:\s*false\s*$'
}
Add-Check 'release.yml: concurrency group: release, cancel-in-progress: false' $concurrencyOk ''

# --- 12. aot-matrix.yml and fuzz.yml -----------------------------------
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
Add-Check 'aot-matrix.yml: matrix, clang, smoke, nightly' ($aotMissing.Count -eq 0) ('missing: ' + ($aotMissing -join ', '))

$fuzzMissing = Find-MissingLiteral -Text $fuzzText -Needles @(
    'cron:',
    'Aid.Xlsx.Fuzzing',
    'timeout 900',
    'continue-on-error: true'
)
Add-Check 'fuzz.yml: target build, short run, no success requirement yet' ($fuzzMissing.Count -eq 0) ('missing: ' + ($fuzzMissing -join ', '))

# --- 13. Directory.Build.props: build + version + metadata --------------
# Version and metadata live in a single Directory.Build.props (the repository owner pattern).
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
Add-Check 'Directory.Build.props: build + version + metadata (TFM net10.0)' ($propsMissing.Count -eq 0) ('missing: ' + ($propsMissing -join ', '))

# The version must not be duplicated in a separate file: the single source is Directory.Build.props.
$versionsFileIssues = @()
if (Test-Path -LiteralPath (Join-Path $RepoRoot 'eng/Versions.props')) {
    $versionsFileIssues += 'eng/Versions.props exists: the version must live only in Directory.Build.props'
}
foreach ($relativePath in ($WorkflowPaths + @('Directory.Build.props', 'Directory.Packages.props'))) {
    $text = Read-RepoFile $relativePath
    if ($null -eq $text) { continue }
    if ($text -match 'Versions\.props') {
        $versionsFileIssues += ($relativePath + ': references the removed Versions.props')
    }
}
Add-Check 'version only in Directory.Build.props (no Versions.props)' ($versionsFileIssues.Count -eq 0) ($versionsFileIssues -join '; ')

# release.yml must read the version from the same place where it is set.
if ($releaseText -notmatch [regex]::Escape('Directory.Build.props')) {
    Add-Check 'release.yml reads the version from Directory.Build.props' $false 'no reference to Directory.Build.props'
}
else {
    Add-Check 'release.yml reads the version from Directory.Build.props' $true
}

# --- 14. .gitattributes: binary corpus ----------------------------------
$gitAttributesText = Read-RepoFile '.gitattributes'
$gitAttributesIssues = @()
foreach ($extension in @('*.xlsx', '*.xlsm', '*.png', '*.bin')) {
    $pattern = '(?m)^' + [regex]::Escape($extension) + '\s+binary\b'
    if ($gitAttributesText -notmatch $pattern) {
        $gitAttributesIssues += ('no line ' + $extension + ' binary')
    }
}
Add-Check '.gitattributes: binary corpus extensions' ($gitAttributesIssues.Count -eq 0) ($gitAttributesIssues -join '; ')

# --- 15. .editorconfig ------------------------------------------------------
$editorConfigText = Read-RepoFile '.editorconfig'
$editorConfigIssues = @()
if ($editorConfigText -notmatch '(?m)^root\s*=\s*true\s*$') { $editorConfigIssues += 'no root = true' }
if ($editorConfigText -notmatch '(?m)^end_of_line\s*=\s*lf\s*$') { $editorConfigIssues += 'no end_of_line = lf' }
if ($editorConfigText -notmatch '\[\*\.cs\]|\[\*\]') { $editorConfigIssues += 'no section for *' }
Add-Check '.editorconfig: root, single line endings' ($editorConfigIssues.Count -eq 0) ($editorConfigIssues -join '; ')

# --- 16. No replace calls in workflows and eng scripts ------------------
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
Add-Check 'no replace calls in workflows/scripts' ($replaceHits.Count -eq 0) ($replaceHits -join ', ')

# --- 17. Trusted Publishing: the main path needs no secrets -------------
# OIDC publishing needs neither an API key nor tokens: only the TP policy on nuget.org
# and the nuget-production environment. The API key is an optional fallback.
$secretsIssues = @()
$releaseLinesForSecrets = Get-Lines $releaseText
$trustedIndex = -1
for ($index = 0; $index -lt $releaseLinesForSecrets.Count; $index++) {
    if ($releaseLinesForSecrets[$index] -match '--api-key\s+"\$TRUSTED_API_KEY"') { $trustedIndex = $index }
}
if ($trustedIndex -lt 0) {
    $secretsIssues += 'no push with --api-key "$TRUSTED_API_KEY" found'
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
            $secretsIssues += ('TP step touches secrets (line ' + ($index + 1) + ')')
        }
    }
}

$fallbackIndex = -1
for ($index = 0; $index -lt $releaseLinesForSecrets.Count; $index++) {
    if ($releaseLinesForSecrets[$index] -match 'secrets\.NUGET_API_KEY') { $fallbackIndex = $index }
}
if ($fallbackIndex -lt 0) {
    $secretsIssues += 'no optional fallback step with secrets.NUGET_API_KEY'
}
else {
    $fallbackEnd = [Math]::Min($releaseLinesForSecrets.Count - 1, $fallbackIndex + 20)
    $fallbackBlock = ($releaseLinesForSecrets[$fallbackIndex..$fallbackEnd] -join "`n")
    if ($fallbackBlock -notmatch 'exit 0') {
        $secretsIssues += 'the fallback is not optional: without the secret there is no exit 0'
    }
}
Add-Check 'TP path without secrets; API key is an optional fallback' ($secretsIssues.Count -eq 0) ($secretsIssues -join '; ')

# --- 18. Encoding -------------------------------------------------------
# Windows PowerShell 5.1 reads BOM-less non-ASCII as ANSI, so props/ps1 files
# with non-ASCII text must be UTF-8 with BOM. Files that are not read by
# PowerShell (.editorconfig, .gitattributes) stay pure ASCII: a BOM hurts there.
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
        $bomIssues += ($relativePath + ': non-ASCII without a UTF-8 BOM (PS 5.1 will read it as ANSI)')
    }
}
Add-Check 'encoding: props/ps1 with non-ASCII use a UTF-8 BOM' ($bomIssues.Count -eq 0) ($bomIssues -join '; ')

$asciiIssues = @()
foreach ($relativePath in @('.editorconfig', '.gitattributes')) {
    $fullPath = Join-Path $RepoRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath)) { continue }
    foreach ($value in ([System.IO.File]::ReadAllBytes($fullPath))) {
        if ($value -gt 127) {
            $asciiIssues += ($relativePath + ': non-ASCII bytes (the file must be ASCII)')
            break
        }
    }
}
Add-Check '.editorconfig/.gitattributes are ASCII without BOM' ($asciiIssues.Count -eq 0) ($asciiIssues -join '; ')

# --- 19. No references to planning documents in committable files -------
# The repository is public: committable files must not reference internal
# planning documents, research files or their section numbers.
# Markers are assembled from fragments (the section sign comes from its code point)
# so that this script does not find the markers in itself and does not false-positive.
#
# References that use our plan section numbering (a section sign followed by 1-2 digits)
# are caught by regexes; specification citations such as ECMA-376 Part 1 section 18.8.30
# are deliberately NOT caught: they are legitimate upstream references, not plan links.
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
    # No git fallback: the known set of files that will land in the repository.
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
            $planningHits += ('planning document reference in a public file: ' + $relativePath + ':' + $lineNumber)
        }
    }
}
# NOT SCANNED: .git/info/exclude and local planning documents: they are not committed
# (excluded via .git/info/exclude), so git ls-files does not return them.
Add-Check 'no references to planning documents in committable files' ($planningHits.Count -eq 0) ($planningHits -join '; ')

# --- 20. C# file names --------------------------------------------------
$csharpNameIssues = @($committableFiles | Where-Object {
    (Test-Path -LiteralPath (Join-Path $RepoRoot $_) -PathType Leaf) -and
        [System.IO.Path]::GetFileName($_) -match '\..*\.cs$' -and
        [System.IO.Path]::GetFileName($_) -notmatch '^[^.]+\.g\.cs$'
})
Add-Check 'C# filenames use one dot or the .g.cs suffix' ($csharpNameIssues.Count -eq 0) ($csharpNameIssues -join '; ')

# --- Summary ------------------------------------------------------------
Write-Host ''
if ($script:Failures.Count -eq 0) {
    Write-Host ('OK: ' + $script:Passed + ' checks passed') -ForegroundColor Green
    exit 0
}

Write-Host ('FAIL: ' + $script:Failures.Count + ' of ' + ($script:Passed + $script:Failures.Count) + ' checks failed') -ForegroundColor Red
foreach ($failure in $script:Failures) {
    Write-Host ('  - ' + $failure) -ForegroundColor Red
}
exit 1
