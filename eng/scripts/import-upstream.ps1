#Requires -Version 5.1
<#
.SYNOPSIS
    Imports the ClosedXML sources at a pinned commit into Aid.Xlsx without upstream history.

.DESCRIPTION
    The import works as follows:
      1. the temporary directory is created OUTSIDE the working copy ($env:TEMP\aidxlsx-import);
      2. exactly one pinned commit is fetched from upstream (fetch --depth 1 --no-tags by SHA);
      3. the tree is exported with `git archive` (a snapshot, no .git and no history);
      4. the ClosedXML, ClosedXML.IO, ClosedXML.IO.CodeGen and ClosedXML.Tests directories
         are copied into the repository (a missing directory is an error);
      5. LICENSE is copied from the snapshot to the repository root if the file is not there yet;
      6. the script verifies that no .git arrived and that the receiving repository history did not change.

    ClosedXML.Parser is NOT vendored by default: the upstream project references the
    ClosedXML.Parser NuGet package, so the build does not need its sources. Vendoring is
    a separate step: the -IncludeParser switch fetches the parser from its own repository
    and puts it into src\ClosedXML.Parser (where it must be retargeted and adapted to the shared props).

    The script needs network access: run it from the repository owner terminal or in CI.
    Without HTTPS access to upstream the script fails with a clear error and exit code 2.

.PARAMETER UpstreamUrl
    Upstream repository URL. Default: https://github.com/ClosedXML/ClosedXML.git

.PARAMETER CommitSha
    Pinned upstream commit. Default: 057f5976c788d97e8eac86d1a528bd5411d9386e

.PARAMETER ParserUpstreamUrl
    ClosedXML.Parser repository URL (used only with -IncludeParser).

.PARAMETER ParserCommitSha
    Pinned ClosedXML.Parser commit (tag 2.0.0-preview1).

.PARAMETER RepoRoot
    Root of the receiving repository. Default: two levels above eng/scripts.

.PARAMETER WorkRoot
    Temporary import directory. Default: $env:TEMP\aidxlsx-import (outside the working copy).

.PARAMETER SnapshotPath
    Offline mode: use an already extracted upstream snapshot instead of the network (for
    script and CI checks without GitHub access).

.PARAMETER ParserSnapshotPath
    Offline mode for the parser when -IncludeParser is set: a ready ClosedXML.Parser snapshot.

.PARAMETER IncludeParser
    Additionally vendor ClosedXML.Parser into src\ClosedXML.Parser (a step for a future phase).

.PARAMETER Force
    Overwrite non-empty destination directories (for example after a failed run).

.PARAMETER CloneFallback
    Fallback: `git clone --filter=blob:none --no-checkout` + `checkout --detach <sha>`
    for servers that reject fetching by SHA.

.PARAMETER DryRun
    Download and copy nothing: check the environment and print the import table.

.PARAMETER KeepWorkDir
    Keep the temporary directory after the run (for incident investigation).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1 -DryRun

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1 -IncludeParser

.NOTES
    Exit codes: 0 - success; 2 - no access to upstream (network/HTTPS); 1 - other errors.
    On some hosts the execution policy blocks unsigned .ps1 files: use the
    -ExecutionPolicy Bypass variant (an environment limitation, not a script problem).
#>
[CmdletBinding()]
param(
    [string] $UpstreamUrl       = 'https://github.com/ClosedXML/ClosedXML.git',
    [string] $CommitSha         = '057f5976c788d97e8eac86d1a528bd5411d9386e',
    [string] $ParserUpstreamUrl = 'https://github.com/ClosedXML/ClosedXML.Parser.git',
    [string] $ParserCommitSha   = '3dcf1fde5741a4cbfbfda8e6142e886f6755463f',
    [string] $RepoRoot          = '',
    [string] $WorkRoot          = '',
    [string] $SnapshotPath      = '',
    [string] $ParserSnapshotPath = '',
    [switch] $IncludeParser,
    [switch] $Force,
    [switch] $CloneFallback,
    [switch] $DryRun,
    [switch] $KeepWorkDir
)

$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}
if (-not $WorkRoot) {
    $WorkRoot = Join-Path $env:TEMP 'aidxlsx-import'
}

# Main payload directories: upstream directory -> Aid.Xlsx directory.
# The names deliberately match upstream: the first import is a mechanical snapshot,
# renaming to Aid.Xlsx.* is done by a separate step.
$MainMap = @(
    [pscustomobject]@{ Source = 'ClosedXML';            Destination = 'src\ClosedXML' },
    [pscustomobject]@{ Source = 'ClosedXML.IO';         Destination = 'src\ClosedXML.IO' },
    [pscustomobject]@{ Source = 'ClosedXML.IO.CodeGen'; Destination = 'src\ClosedXML.IO.CodeGen' },
    [pscustomobject]@{ Source = 'ClosedXML.Tests';      Destination = 'tests\ClosedXML.Tests' }
)

# Parser vendoring (only with -IncludeParser): inside the ClosedXML.Parser repository
# the project lives in src/ClosedXML.Parser and the license lives in the repository root.
$ParserProjectSource      = 'src/ClosedXML.Parser'
$ParserProjectDestination = 'src\ClosedXML.Parser'
$ParserLicenseFileName    = 'LICENSE.txt'

$script:ImportRows = @()

function Write-Step {
    param([string] $Message)
    Write-Host ('[import] ' + $Message)
}

function Add-ImportRow {
    param(
        [string] $Target,
        [string] $Destination,
        [string] $Status
    )
    $script:ImportRows += [pscustomobject]@{ Target = $Target; Destination = $Destination; Status = $Status }
}

function Write-ImportTable {
    $rows = $script:ImportRows
    if (($null -eq $rows) -or ($rows.Count -eq 0)) { return }

    $targetWidth = 8
    $destinationWidth = 11
    $statusWidth = 6
    foreach ($row in $rows) {
        if ($row.Target.Length -gt $targetWidth) { $targetWidth = $row.Target.Length }
        if ($row.Destination.Length -gt $destinationWidth) { $destinationWidth = $row.Destination.Length }
        if ($row.Status.Length -gt $statusWidth) { $statusWidth = $row.Status.Length }
    }

    $format = '{0,-' + $targetWidth + '}  {1,-' + $destinationWidth + '}  {2}'
    Write-Host ''
    Write-Host '  Import upstream -> Aid.Xlsx'
    Write-Host ('  ' + ('-' * ($targetWidth + $destinationWidth + $statusWidth + 4)))
    Write-Host ('  ' + ($format -f 'upstream', 'Aid.Xlsx', 'status'))
    Write-Host ('  ' + ('-' * ($targetWidth + $destinationWidth + $statusWidth + 4)))
    foreach ($row in $rows) {
        Write-Host ('  ' + ($format -f $row.Target, $row.Destination, $row.Status))
    }
    Write-Host ''
}

function Get-UpstreamSnapshot {
    param(
        [string] $Url,
        [string] $Sha,
        [string] $Label,
        [string] $SnapshotWorkDir,
        [string] $ProvidedPath,
        [switch] $UseCloneFallback
    )

    if ($ProvidedPath) {
        if (-not (Test-Path -LiteralPath $ProvidedPath)) {
            throw ('Snapshot ' + $Label + ' not found: ' + $ProvidedPath)
        }
        $resolved = (Resolve-Path -LiteralPath $ProvidedPath).Path
        Write-Step ($Label + ': using a ready snapshot ' + $resolved + ' (network not needed)')
        return $resolved
    }

    $scratch = Join-Path $SnapshotWorkDir 'upstream'
    $archive = Join-Path $SnapshotWorkDir ('snapshot-' + $Sha + '.zip')
    $snapshot = Join-Path $SnapshotWorkDir 'tree'

    if (Test-Path -LiteralPath $SnapshotWorkDir) {
        Remove-Item -LiteralPath $SnapshotWorkDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $SnapshotWorkDir -Force | Out-Null

    if ($UseCloneFallback) {
        Write-Step ($Label + ': cloning without history (--filter=blob:none --no-checkout)')
        $clone = Invoke-Git -Arguments @('clone', '--quiet', '--filter=blob:none', '--no-checkout', $Url, $scratch)
        if ($clone.ExitCode -ne 0) {
            $net = Find-NetworkFailure $clone.Output
            if ($net) {
                throw ("No access to upstream ($Url): matched signature '$net'. " +
                       'Outbound HTTPS may be blocked; run the script from the repository owner terminal or in CI.')
            }
            throw ('git clone failed. git output: ' + $clone.Output)
        }
        $checkout = Invoke-Git -Arguments @('-C', $scratch, 'checkout', '--quiet', '--detach', $Sha)
        if ($checkout.ExitCode -ne 0) {
            throw ('Could not check out commit ' + $Sha + '. git output: ' + $checkout.Output)
        }
    }
    else {
        Write-Step ($Label + ': initializing the temporary repository')
        $init = Invoke-Git -Arguments @('init', '--quiet', $scratch)
        if ($init.ExitCode -ne 0) {
            throw ('git init failed. git output: ' + $init.Output)
        }
        $remote = Invoke-Git -Arguments @('-C', $scratch, 'remote', 'add', 'origin', $Url)
        if ($remote.ExitCode -ne 0) {
            throw ('git remote add failed. git output: ' + $remote.Output)
        }

        Write-Step ($Label + ': fetch ' + $Sha + ' (--depth 1 --no-tags)')
        $fetch = Invoke-Git -Arguments @('-C', $scratch, 'fetch', '--quiet', '--depth', '1', '--no-tags', 'origin', $Sha)
        if ($fetch.ExitCode -ne 0) {
            $net = Find-NetworkFailure $fetch.Output
            if ($net) {
                throw ("No access to upstream ($Url): matched signature '$net'. " +
                       'Outbound HTTPS may be blocked; run the script from the repository owner terminal or in CI.')
            }
            throw ('git fetch could not get commit ' + $Sha + ' from ' + $Url + '. ' +
                   'If the server rejects fetching by SHA, retry with -CloneFallback. git output: ' + $fetch.Output)
        }
    }

    $commitCheck = Invoke-Git -Arguments @('-C', $scratch, 'cat-file', '-e', ($Sha + '^{commit}'))
    if ($commitCheck.ExitCode -ne 0) {
        throw ('Commit ' + $Sha + ' is missing from the fetched ' + $Label + ' repository. ' +
               'Check the SHA: the commit may have been rewritten upstream.')
    }

    $archiveResult = Invoke-Git -Arguments @('-C', $scratch, 'archive', '--format=zip', '--output', $archive, $Sha)
    if ($archiveResult.ExitCode -ne 0) {
        throw ('git archive failed (' + $Label + '). git output: ' + $archiveResult.Output)
    }
    if (-not (Test-Path -LiteralPath $archive)) {
        throw ('git archive did not create the archive (' + $Label + '): ' + $archive)
    }

    New-Item -ItemType Directory -Path $snapshot -Force | Out-Null
    Expand-Archive -LiteralPath $archive -DestinationPath $snapshot -Force
    Write-Step ($Label + ': snapshot extracted, commit verified')

    return $snapshot
}

function Invoke-Git {
    param(
        [string[]] $Arguments,
        [string]   $WorkingDirectory = ''
    )

    $previousLocation = $null
    $previousErrorAction = $ErrorActionPreference
    try {
        # git writes progress and warnings to stderr. In Windows PowerShell 5.1 with
        # $ErrorActionPreference='Stop' the first stderr line becomes a terminating error,
        # so inside the git call stderr is treated as data, not as an error.
        $ErrorActionPreference = 'Continue'
        if ($WorkingDirectory) {
            $previousLocation = Get-Location
            Set-Location -LiteralPath $WorkingDirectory
        }
        $lines = @(& git @Arguments 2>&1 | ForEach-Object { '' + $_ })
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
        if ($null -ne $previousLocation) {
            Set-Location -LiteralPath $previousLocation
        }
    }

    return [pscustomobject]@{
        ExitCode = $code
        Output   = ($lines -join [Environment]::NewLine)
    }
}

function Find-NetworkFailure {
    param([string] $Text)

    $patterns = @(
        'Could not resolve host',
        'unable to access',
        'Failed to connect',
        'Connection refused',
        'Connection timed out',
        'Operation timed out',
        'RPC failed',
        'Could not read from remote repository',
        'Empty reply from server',
        'SSL certificate problem',
        'gnutls',
        'TLS',
        'proxy'
    )

    foreach ($pattern in $patterns) {
        if ($Text -match [regex]::Escape($pattern)) {
            return $pattern
        }
    }
    return ''
}

function Assert-NotInsideRepository {
    param(
        [string] $Candidate,
        [string] $RepositoryRoot
    )

    $candidateFull = [System.IO.Path]::GetFullPath($Candidate)
    $rootFull = [System.IO.Path]::GetFullPath($RepositoryRoot)
    $rootWithSeparator = $rootFull.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar

    if ($candidateFull.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw ("Temporary directory '$candidateFull' is inside the repository '$rootFull'. " +
               'The import must run outside the working copy: pass -WorkRoot pointing elsewhere.')
    }
}

function Copy-DirectoryFromSnapshot {
    param(
        [string] $SnapshotRoot,
        [string] $SourceRelative,
        [string] $DestinationRelative
    )

    $sourcePath = Join-Path $SnapshotRoot $SourceRelative
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        return $false
    }

    $destinationPath = Join-Path $RepoRoot $DestinationRelative
    $parent = Split-Path -Path $destinationPath -Parent
    New-Item -ItemType Directory -Path $parent -Force | Out-Null

    if (Test-Path -LiteralPath $destinationPath) {
        Remove-Item -LiteralPath $destinationPath -Recurse -Force
    }
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Recurse -Force
    return $true
}

function Assert-NoNestedGit {
    param([string] $DestinationRelative)

    $destinationPath = Join-Path $RepoRoot $DestinationRelative
    $nested = @(Get-ChildItem -LiteralPath $destinationPath -Recurse -Force -Directory -Filter '.git' -ErrorAction SilentlyContinue)
    if ($nested.Count -gt 0) {
        throw ('Found .git after copying: ' + $nested[0].FullName + '. The import is corrupted, revert the directory.')
    }
}

try {
    Write-Step ('repository: ' + $RepoRoot)
    Write-Step ('upstream:    ' + $UpstreamUrl + ' @ ' + $CommitSha)
    if ($IncludeParser) {
        Write-Step ('parser:      ' + $ParserUpstreamUrl + ' @ ' + $ParserCommitSha + ' (-IncludeParser)')
    }
    Write-Step ('temp dir:   ' + $WorkRoot)

    # --- 1. Environment -----------------------------------------------------
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCommand) {
        throw 'git was not found in PATH. Install Git for Windows and retry.'
    }

    if (-not (Test-Path -LiteralPath $RepoRoot)) {
        throw ("Repository root not found: '$RepoRoot'.")
    }

    $inside = Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-parse', '--is-inside-work-tree')
    if ($inside.ExitCode -ne 0 -or ($inside.Output.Trim() -ne 'true')) {
        throw ("Directory '$RepoRoot' is not a git repository. " +
               "Run 'git init' or pass -RepoRoot. git output: " + $inside.Output)
    }

    Assert-NotInsideRepository -Candidate $WorkRoot -RepositoryRoot $RepoRoot
    Write-Step 'temporary directory is outside the working copy: OK'

    $commitsBefore = (Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-list', '--all', '--count')).Output.Trim()
    if (-not $commitsBefore) {
        throw 'Could not read the receiving repository history (git rev-list --all --count).'
    }

    # --- 2. Upstream snapshot -----------------------------------------------
    $mainSnapshotWorkDir = Join-Path $WorkRoot 'main'
    $mainSnapshot = Get-UpstreamSnapshot -Url $UpstreamUrl -Sha $CommitSha -Label 'closedxml' `
        -SnapshotWorkDir $mainSnapshotWorkDir -ProvidedPath $SnapshotPath -UseCloneFallback:$CloneFallback

    # --- 3. Snapshot contents: table before copying --------------------------
    $missing = @()
    foreach ($item in $MainMap) {
        $sourcePath = Join-Path $mainSnapshot $item.Source
        if (Test-Path -LiteralPath $sourcePath) {
            Add-ImportRow -Target $item.Source -Destination $item.Destination -Status 'present in snapshot'
        }
        else {
            Add-ImportRow -Target $item.Source -Destination $item.Destination -Status 'MISSING in snapshot'
            $missing += $item.Source
        }
    }

    if ($missing.Count -gt 0) {
        Write-ImportTable
        throw ('Snapshot of commit ' + $CommitSha + ' has no directories: ' + ($missing -join ', ') + '. ' +
               'Upstream may have changed its layout: compare with the pinned commit and update the import list.')
    }

    $licenseSource = Join-Path $mainSnapshot 'LICENSE'
    $licenseTarget = Join-Path $RepoRoot 'LICENSE'

    # --- 4. Destination check (the error is raised after DryRun) -------------
    $nonEmpty = @()
    foreach ($item in $MainMap) {
        $destinationPath = Join-Path $RepoRoot $item.Destination
        if (Test-Path -LiteralPath $destinationPath) {
            $children = @(Get-ChildItem -LiteralPath $destinationPath -Force -ErrorAction SilentlyContinue)
            if ($children.Count -gt 0) {
                $nonEmpty += $item.Destination
            }
        }
    }

    # --- 5. DryRun: print the plan only --------------------------------------
    if ($DryRun) {
        Write-Step '-DryRun mode: the network is not used, nothing is copied.'

        foreach ($row in $script:ImportRows) {
            if ($row.Status -eq 'present in snapshot') {
                if ($nonEmpty -contains $row.Destination) {
                    $row.Status = 'plan: overwrite (-Force)'
                }
                else {
                    $row.Status = 'plan: copy'
                }
            }
        }

        if (Test-Path -LiteralPath $licenseTarget) {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'plan: already exists'
        }
        elseif (Test-Path -LiteralPath $licenseSource) {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'plan: copy to root'
        }
        else {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'plan: missing in snapshot'
        }

        if ($IncludeParser) {
            if ($ParserSnapshotPath) {
                $parserProject = Join-Path $ParserSnapshotPath $ParserProjectSource
                if (Test-Path -LiteralPath $parserProject) {
                    Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'plan: copy'
                }
                else {
                    Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'plan: MISSING in snapshot'
                }
            }
            else {
                Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'plan: copy (network)'
            }
        }
        else {
            Add-ImportRow -Target 'ClosedXML.Parser' -Destination $ParserProjectDestination -Status 'skipped: NuGet package'
        }

        Write-ImportTable
        if ($nonEmpty.Count -gt 0) {
            Write-Step ('occupied directories: ' + ($nonEmpty -join ', ') + '. A real run needs -Force.')
        }
        if (-not $IncludeParser) {
            Write-Step 'Parser is not vendored by default (the NuGet package is used); vendoring is a separate -IncludeParser step.'
        }
        Write-Host '[import] OK (dry-run)'
        exit 0
    }

    if ($nonEmpty.Count -gt 0 -and (-not $Force)) {
        Write-ImportTable
        throw ('Directories already exist and are not empty: ' + ($nonEmpty -join ', ') + '. ' +
               'This is normal after a failed run: retry with -Force (the import overwrites them from the snapshot).')
    }

    # --- 6. Copy the main payload --------------------------------------------
    foreach ($item in $MainMap) {
        if (-not (Copy-DirectoryFromSnapshot -SnapshotRoot $mainSnapshot -SourceRelative $item.Source -DestinationRelative $item.Destination)) {
            throw ('Snapshot of commit ' + $CommitSha + ' has no directory ' + $item.Source + '.')
        }
        foreach ($row in $script:ImportRows) {
            if (($row.Target -eq $item.Source) -and ($row.Destination -eq $item.Destination)) {
                $row.Status = 'copied'
            }
        }
    }

    # --- 7. LICENSE from the snapshot to the root (if not there yet) ----------
    if (Test-Path -LiteralPath $licenseTarget) {
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'already exists (left untouched)'
    }
    elseif (Test-Path -LiteralPath $licenseSource) {
        Copy-Item -LiteralPath $licenseSource -Destination $licenseTarget -Force
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'copied to root'
    }
    else {
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'missing in snapshot (check manually)'
    }

    # --- 8. Parser (only with -IncludeParser) --------------------------------
    if ($IncludeParser) {
        $parserSnapshotWorkDir = Join-Path $WorkRoot 'parser'
        $parserSnapshot = Get-UpstreamSnapshot -Url $ParserUpstreamUrl -Sha $ParserCommitSha -Label 'parser' `
            -SnapshotWorkDir $parserSnapshotWorkDir -ProvidedPath $ParserSnapshotPath -UseCloneFallback:$CloneFallback

        if (-not (Copy-DirectoryFromSnapshot -SnapshotRoot $parserSnapshot -SourceRelative $ParserProjectSource -DestinationRelative $ParserProjectDestination)) {
            throw ('Snapshot of the ClosedXML.Parser repository has no directory ' + $ParserProjectSource + '.')
        }
        Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'copied'

        $parserLicenseSource = Join-Path $parserSnapshot $ParserLicenseFileName
        $parserLicenseTarget = Join-Path (Join-Path $RepoRoot $ParserProjectDestination) $ParserLicenseFileName
        if ((Test-Path -LiteralPath $parserLicenseSource) -and (-not (Test-Path -LiteralPath $parserLicenseTarget))) {
            New-Item -ItemType Directory -Path (Split-Path -Path $parserLicenseTarget -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $parserLicenseSource -Destination $parserLicenseTarget -Force
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'copied'
        }
        elseif (Test-Path -LiteralPath $parserLicenseTarget) {
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'already exists'
        }
        else {
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'missing in snapshot'
        }
    }
    else {
        Add-ImportRow -Target 'ClosedXML.Parser' -Destination $ParserProjectDestination -Status 'skipped: NuGet package'
    }

    # --- 9. Post-checks ------------------------------------------------------
    foreach ($item in $MainMap) {
        Assert-NoNestedGit -DestinationRelative $item.Destination
    }
    if ($IncludeParser) {
        Assert-NoNestedGit -DestinationRelative $ParserProjectDestination
    }
    Write-Step 'check "no .git": OK'

    $commitsAfter = (Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-list', '--all', '--count')).Output.Trim()
    if ($commitsAfter -ne $commitsBefore) {
        throw ('The receiving repository history changed (' + $commitsBefore + ' -> ' + $commitsAfter + '). ' +
               'The import must not touch history: check `git remote -v` and `git log`.')
    }
    Write-Step ('check "no foreign history": OK (' + $commitsAfter + ' commits)')

    $remotes = (Invoke-Git -Arguments @('-C', $RepoRoot, 'remote', '-v')).Output
    foreach ($line in ($remotes -split "`n")) {
        if ($line -match 'closedxml') {
            throw ('The repository gained a remote pointing at upstream: ' + $line.Trim() + '. Remove it with: git remote remove <name>.')
        }
    }
    Write-Step 'remotes check: OK'

    Write-ImportTable
    if (-not $IncludeParser) {
        Write-Step 'Parser is not vendored by default (the NuGet package is used); vendoring is a separate -IncludeParser step.'
    }
    Write-Step 'done. Sources are in src/ and tests/, history is untouched.'
    Write-Step 'next: git add -A; git commit -m "chore: bootstrap Aid.Xlsx solution"'
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-Host ''
    Write-Host ('[import] FAIL: ' + $message) -ForegroundColor Red
    if ($message -like '*No access to upstream*') {
        exit 2
    }
    exit 1
}
