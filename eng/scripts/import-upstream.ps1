#Requires -Version 5.1
<#
.SYNOPSIS
    Импортирует исходники ClosedXML на pinned-коммите в Aid.Xlsx без upstream-истории.

.DESCRIPTION
    Импорт устроен так:
      1. временный каталог создаётся ВНЕ рабочей копии ($env:TEMP\aidxlsx-import);
      2. из upstream забирается ровно один pinned-коммит (fetch --depth 1 --no-tags по SHA);
      3. дерево выгружается через `git archive` (снимок, без .git и без истории);
      4. в репозиторий копируются каталоги ClosedXML, ClosedXML.IO, ClosedXML.IO.CodeGen
         и ClosedXML.Tests (их отсутствие в снимке — ошибка);
      5. LICENSE из снимка копируется в корень репозитория, если файла там ещё нет;
      6. проверяется, что не приехал .git и что история принимающего репозитория не изменилась.

    ClosedXML.Parser в фазе 0 по умолчанию НЕ вендорится: upstream-проект ссылается на
    NuGet-пакет ClosedXML.Parser, поэтому сборка не требует его исходников. Вендоринг —
    отдельный шаг: switch -IncludeParser забирает parser из его собственного репозитория
    и кладёт в src\ClosedXML.Parser (там его нужно ретаргетить и адаптировать под общие props).

    Скрипт сетевой: запускать в обычном терминале владельца или в CI.
    Без доступа к upstream по HTTPS скрипт завершится с понятной ошибкой и кодом 2.

.PARAMETER UpstreamUrl
    URL upstream-репозитория. По умолчанию https://github.com/ClosedXML/ClosedXML.git

.PARAMETER CommitSha
    Pinned-коммит upstream. По умолчанию 057f5976c788d97e8eac86d1a528bd5411d9386e

.PARAMETER ParserUpstreamUrl
    URL репозитория ClosedXML.Parser (используется только с -IncludeParser).

.PARAMETER ParserCommitSha
    Pinned-коммит ClosedXML.Parser (тег 2.0.0-preview1).

.PARAMETER RepoRoot
    Корень принимающего репозитория. По умолчанию — два уровня выше eng/scripts.

.PARAMETER WorkRoot
    Временный каталог импорта. По умолчанию $env:TEMP\aidxlsx-import (вне рабочей копии).

.PARAMETER SnapshotPath
    Офлайн-режим: использовать уже развёрнутый снимок upstream вместо сети (для проверки
    скрипта и CI без доступа к GitHub).

.PARAMETER ParserSnapshotPath
    Офлайн-режим для parser при -IncludeParser: готовый снимок репозитория ClosedXML.Parser.

.PARAMETER IncludeParser
    Дополнительно вендорить ClosedXML.Parser в src\ClosedXML.Parser (шаг для будущей фазы).

.PARAMETER Force
    Перезаписать непустые каталоги назначения (например, после неудачного прогона).

.PARAMETER CloneFallback
    Резервный путь: `git clone --filter=blob:none --no-checkout` + `checkout --detach <sha>`
    для серверов, запрещающих fetch по SHA.

.PARAMETER DryRun
    Ничего не скачивать и не копировать: проверить окружение, показать таблицу импорта.

.PARAMETER KeepWorkDir
    Не удалять временный каталог после завершения (для разбора инцидентов).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1 -DryRun

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\eng\scripts\import-upstream.ps1 -IncludeParser

.NOTES
    Коды выхода: 0 — успех; 2 — нет доступа к upstream (сеть/HTTPS); 1 — прочие ошибки.
    На части хостов политика выполнения запрещает неподписанные .ps1 — используйте
    -ExecutionPolicy Bypass (это ограничение окружения, а не скрипта).
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

# Каталоги основной поставки: каталог upstream -> каталог Aid.Xlsx.
# Имена намеренно совпадают с upstream: первый импорт — механический снимок,
# переименование в Aid.Xlsx.* делается отдельным шагом.
$MainMap = @(
    [pscustomobject]@{ Source = 'ClosedXML';            Destination = 'src\ClosedXML' },
    [pscustomobject]@{ Source = 'ClosedXML.IO';         Destination = 'src\ClosedXML.IO' },
    [pscustomobject]@{ Source = 'ClosedXML.IO.CodeGen'; Destination = 'src\ClosedXML.IO.CodeGen' },
    [pscustomobject]@{ Source = 'ClosedXML.Tests';      Destination = 'tests\ClosedXML.Tests' }
)

# Вендоринг parser (только с -IncludeParser): внутри репозитория ClosedXML.Parser
# проект лежит в src/ClosedXML.Parser, лицензия — в корне репозитория.
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
    Write-Host '  Импорт upstream -> Aid.Xlsx'
    Write-Host ('  ' + ('-' * ($targetWidth + $destinationWidth + $statusWidth + 4)))
    Write-Host ('  ' + ($format -f 'upstream', 'Aid.Xlsx', 'статус'))
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
            throw ('Снимок ' + $Label + ' не найден: ' + $ProvidedPath)
        }
        $resolved = (Resolve-Path -LiteralPath $ProvidedPath).Path
        Write-Step ($Label + ': используется готовый снимок ' + $resolved + ' (сеть не нужна)')
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
        Write-Step ($Label + ': клонирование без истории (--filter=blob:none --no-checkout)')
        $clone = Invoke-Git -Arguments @('clone', '--quiet', '--filter=blob:none', '--no-checkout', $Url, $scratch)
        if ($clone.ExitCode -ne 0) {
            $net = Find-NetworkFailure $clone.Output
            if ($net) {
                throw ("Нет доступа к upstream ($Url): совпал признак '$net'. " +
                       'Исходящий HTTPS может быть заблокирован; запустите скрипт в обычном терминале владельца или в CI.')
            }
            throw ('git clone не удался. Вывод git: ' + $clone.Output)
        }
        $checkout = Invoke-Git -Arguments @('-C', $scratch, 'checkout', '--quiet', '--detach', $Sha)
        if ($checkout.ExitCode -ne 0) {
            throw ('Не удалось переключиться на коммит ' + $Sha + '. Вывод git: ' + $checkout.Output)
        }
    }
    else {
        Write-Step ($Label + ': инициализация временного репозитория')
        $init = Invoke-Git -Arguments @('init', '--quiet', $scratch)
        if ($init.ExitCode -ne 0) {
            throw ('git init не удался. Вывод git: ' + $init.Output)
        }
        $remote = Invoke-Git -Arguments @('-C', $scratch, 'remote', 'add', 'origin', $Url)
        if ($remote.ExitCode -ne 0) {
            throw ('git remote add не удался. Вывод git: ' + $remote.Output)
        }

        Write-Step ($Label + ': fetch ' + $Sha + ' (--depth 1 --no-tags)')
        $fetch = Invoke-Git -Arguments @('-C', $scratch, 'fetch', '--quiet', '--depth', '1', '--no-tags', 'origin', $Sha)
        if ($fetch.ExitCode -ne 0) {
            $net = Find-NetworkFailure $fetch.Output
            if ($net) {
                throw ("Нет доступа к upstream ($Url): совпал признак '$net'. " +
                       'Исходящий HTTPS может быть заблокирован; запустите скрипт в обычном терминале владельца или в CI.')
            }
            throw ('git fetch не смог получить коммит ' + $Sha + ' из ' + $Url + '. ' +
                   'Если сервер запрещает fetch по SHA, повторите с -CloneFallback. Вывод git: ' + $fetch.Output)
        }
    }

    $commitCheck = Invoke-Git -Arguments @('-C', $scratch, 'cat-file', '-e', ($Sha + '^{commit}'))
    if ($commitCheck.ExitCode -ne 0) {
        throw ('Коммит ' + $Sha + ' отсутствует в полученном репозитории ' + $Label + '. ' +
               'Проверьте SHA: возможно, коммит переписан в upstream.')
    }

    $archiveResult = Invoke-Git -Arguments @('-C', $scratch, 'archive', '--format=zip', '--output', $archive, $Sha)
    if ($archiveResult.ExitCode -ne 0) {
        throw ('git archive не удался (' + $Label + '). Вывод git: ' + $archiveResult.Output)
    }
    if (-not (Test-Path -LiteralPath $archive)) {
        throw ('git archive не создал архив (' + $Label + '): ' + $archive)
    }

    New-Item -ItemType Directory -Path $snapshot -Force | Out-Null
    Expand-Archive -LiteralPath $archive -DestinationPath $snapshot -Force
    Write-Step ($Label + ': снимок развёрнут, коммит подтверждён')

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
        # git пишет прогресс и предупреждения в stderr. В Windows PowerShell 5.1 при
        # $ErrorActionPreference='Stop' первая строка stderr превращается в terminating error,
        # поэтому внутри вызова git stderr считается данными, а не ошибкой.
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
        throw ("Временный каталог '$candidateFull' находится внутри репозитория '$rootFull'. " +
               'Импорт обязан идти вне рабочей копии — укажите -WorkRoot снаружи.')
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
        throw ('После копирования найден .git: ' + $nested[0].FullName + ' — импорт испорчен, откатите каталог.')
    }
}

try {
    Write-Step ('репозиторий: ' + $RepoRoot)
    Write-Step ('upstream:    ' + $UpstreamUrl + ' @ ' + $CommitSha)
    if ($IncludeParser) {
        Write-Step ('parser:      ' + $ParserUpstreamUrl + ' @ ' + $ParserCommitSha + ' (-IncludeParser)')
    }
    Write-Step ('временный:   ' + $WorkRoot)

    # --- 1. Окружение -------------------------------------------------------
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCommand) {
        throw 'git не найден в PATH. Установите Git for Windows и повторите.'
    }

    if (-not (Test-Path -LiteralPath $RepoRoot)) {
        throw ("Корень репозитория не найден: '$RepoRoot'.")
    }

    $inside = Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-parse', '--is-inside-work-tree')
    if ($inside.ExitCode -ne 0 -or ($inside.Output.Trim() -ne 'true')) {
        throw ("Каталог '$RepoRoot' не является git-репозиторием. " +
               "Выполните 'git init' или укажите -RepoRoot. Вывод git: " + $inside.Output)
    }

    Assert-NotInsideRepository -Candidate $WorkRoot -RepositoryRoot $RepoRoot
    Write-Step 'временный каталог вне рабочей копии: OK'

    $commitsBefore = (Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-list', '--all', '--count')).Output.Trim()
    if (-not $commitsBefore) {
        throw 'Не удалось прочитать историю принимающего репозитория (git rev-list --all --count).'
    }

    # --- 2. Снимок upstream -------------------------------------------------
    $mainSnapshotWorkDir = Join-Path $WorkRoot 'main'
    $mainSnapshot = Get-UpstreamSnapshot -Url $UpstreamUrl -Sha $CommitSha -Label 'closedxml' `
        -SnapshotWorkDir $mainSnapshotWorkDir -ProvidedPath $SnapshotPath -UseCloneFallback:$CloneFallback

    # --- 3. Что есть в снимке: таблица до копирования ------------------------
    $missing = @()
    foreach ($item in $MainMap) {
        $sourcePath = Join-Path $mainSnapshot $item.Source
        if (Test-Path -LiteralPath $sourcePath) {
            Add-ImportRow -Target $item.Source -Destination $item.Destination -Status 'есть в снимке'
        }
        else {
            Add-ImportRow -Target $item.Source -Destination $item.Destination -Status 'НЕТ в снимке'
            $missing += $item.Source
        }
    }

    if ($missing.Count -gt 0) {
        Write-ImportTable
        throw ('В снимке коммита ' + $CommitSha + ' нет каталогов: ' + ($missing -join ', ') + '. ' +
               'Upstream мог изменить структуру — сверьтесь с pinned-коммитом и обновите список импорта.')
    }

    $licenseSource = Join-Path $mainSnapshot 'LICENSE'
    $licenseTarget = Join-Path $RepoRoot 'LICENSE'

    # --- 4. Проверка занятых каталогов (ошибка — после DryRun) --------------
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

    # --- 5. DryRun: только показать план ------------------------------------
    if ($DryRun) {
        Write-Step 'режим -DryRun: сеть не используется, ничего не копируется.'

        foreach ($row in $script:ImportRows) {
            if ($row.Status -eq 'есть в снимке') {
                if ($nonEmpty -contains $row.Destination) {
                    $row.Status = 'план: перезаписать (-Force)'
                }
                else {
                    $row.Status = 'план: скопировать'
                }
            }
        }

        if (Test-Path -LiteralPath $licenseTarget) {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'план: уже существует'
        }
        elseif (Test-Path -LiteralPath $licenseSource) {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'план: скопировать в корень'
        }
        else {
            Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'план: нет в снимке'
        }

        if ($IncludeParser) {
            if ($ParserSnapshotPath) {
                $parserProject = Join-Path $ParserSnapshotPath $ParserProjectSource
                if (Test-Path -LiteralPath $parserProject) {
                    Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'план: скопировать'
                }
                else {
                    Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'план: НЕТ в снимке'
                }
            }
            else {
                Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'план: скопировать (сеть)'
            }
        }
        else {
            Add-ImportRow -Target 'ClosedXML.Parser' -Destination $ParserProjectDestination -Status 'пропущен: NuGet-пакет'
        }

        Write-ImportTable
        if ($nonEmpty.Count -gt 0) {
            Write-Step ('занятые каталоги: ' + ($nonEmpty -join ', ') + ' — для реального прогона нужен -Force.')
        }
        if (-not $IncludeParser) {
            Write-Step 'Parser не вендорится в фазе 0 (используется NuGet-пакет); вендоринг — отдельный шаг -IncludeParser.'
        }
        Write-Host '[import] OK (dry-run)'
        exit 0
    }

    if ($nonEmpty.Count -gt 0 -and (-not $Force)) {
        Write-ImportTable
        throw ('Каталоги уже существуют и не пусты: ' + ($nonEmpty -join ', ') + '. ' +
               'Это нормально после неудачного прогона: повторите с -Force (импорт перезапишет их из снимка).')
    }

    # --- 6. Копирование основной поставки ----------------------------------
    foreach ($item in $MainMap) {
        if (-not (Copy-DirectoryFromSnapshot -SnapshotRoot $mainSnapshot -SourceRelative $item.Source -DestinationRelative $item.Destination)) {
            throw ('В снимке коммита ' + $CommitSha + ' нет каталога ' + $item.Source + '.')
        }
        foreach ($row in $script:ImportRows) {
            if (($row.Target -eq $item.Source) -and ($row.Destination -eq $item.Destination)) {
                $row.Status = 'скопирован'
            }
        }
    }

    # --- 7. LICENSE из снимка в корень (если файла ещё нет) -----------------
    if (Test-Path -LiteralPath $licenseTarget) {
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'уже существует (не тронут)'
    }
    elseif (Test-Path -LiteralPath $licenseSource) {
        Copy-Item -LiteralPath $licenseSource -Destination $licenseTarget -Force
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'скопирован в корень'
    }
    else {
        Add-ImportRow -Target 'LICENSE' -Destination 'LICENSE' -Status 'нет в снимке (проверьте вручную)'
    }

    # --- 8. Parser (только с -IncludeParser) --------------------------------
    if ($IncludeParser) {
        $parserSnapshotWorkDir = Join-Path $WorkRoot 'parser'
        $parserSnapshot = Get-UpstreamSnapshot -Url $ParserUpstreamUrl -Sha $ParserCommitSha -Label 'parser' `
            -SnapshotWorkDir $parserSnapshotWorkDir -ProvidedPath $ParserSnapshotPath -UseCloneFallback:$CloneFallback

        if (-not (Copy-DirectoryFromSnapshot -SnapshotRoot $parserSnapshot -SourceRelative $ParserProjectSource -DestinationRelative $ParserProjectDestination)) {
            throw ('В снимке репозитория ClosedXML.Parser нет каталога ' + $ParserProjectSource + '.')
        }
        Add-ImportRow -Target $ParserProjectSource -Destination $ParserProjectDestination -Status 'скопирован'

        $parserLicenseSource = Join-Path $parserSnapshot $ParserLicenseFileName
        $parserLicenseTarget = Join-Path (Join-Path $RepoRoot $ParserProjectDestination) $ParserLicenseFileName
        if ((Test-Path -LiteralPath $parserLicenseSource) -and (-not (Test-Path -LiteralPath $parserLicenseTarget))) {
            New-Item -ItemType Directory -Path (Split-Path -Path $parserLicenseTarget -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $parserLicenseSource -Destination $parserLicenseTarget -Force
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'скопирован'
        }
        elseif (Test-Path -LiteralPath $parserLicenseTarget) {
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'уже существует'
        }
        else {
            Add-ImportRow -Target $ParserLicenseFileName -Destination ($ParserProjectDestination + '\' + $ParserLicenseFileName) -Status 'нет в снимке'
        }
    }
    else {
        Add-ImportRow -Target 'ClosedXML.Parser' -Destination $ParserProjectDestination -Status 'пропущен: NuGet-пакет'
    }

    # --- 9. Пост-проверки ---------------------------------------------------
    foreach ($item in $MainMap) {
        Assert-NoNestedGit -DestinationRelative $item.Destination
    }
    if ($IncludeParser) {
        Assert-NoNestedGit -DestinationRelative $ParserProjectDestination
    }
    Write-Step 'проверка «нет .git»: OK'

    $commitsAfter = (Invoke-Git -Arguments @('-C', $RepoRoot, 'rev-list', '--all', '--count')).Output.Trim()
    if ($commitsAfter -ne $commitsBefore) {
        throw ('История принимающего репозитория изменилась (' + $commitsBefore + ' -> ' + $commitsAfter + '). ' +
               'Импорт не должен трогать историю: проверьте `git remote -v` и `git log`.')
    }
    Write-Step ('проверка «чужая история не приехала»: OK (коммитов ' + $commitsAfter + ')')

    $remotes = (Invoke-Git -Arguments @('-C', $RepoRoot, 'remote', '-v')).Output
    foreach ($line in ($remotes -split "`n")) {
        if ($line -match 'closedxml') {
            throw ('В репозитории появился remote на upstream: ' + $line.Trim() + '. Удалите его: git remote remove <name>.')
        }
    }
    Write-Step 'проверка remotes: OK'

    Write-ImportTable
    if (-not $IncludeParser) {
        Write-Step 'Parser не вендорится в фазе 0 (используется NuGet-пакет); вендоринг — отдельный шаг -IncludeParser.'
    }
    Write-Step 'готово. Исходники в src/ и tests/, история не тронута.'
    Write-Step 'дальше: git add -A; git commit -m "chore: bootstrap Aid.Xlsx solution"'
    exit 0
}
catch {
    $message = $_.Exception.Message
    Write-Host ''
    Write-Host ('[import] FAIL: ' + $message) -ForegroundColor Red
    if ($message -like '*Нет доступа к upstream*') {
        exit 2
    }
    exit 1
}