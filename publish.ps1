#Requires -Version 7.0
<#
.SYNOPSIS
    Builds, tests, packs and shape-checks Avig.Ord.

.DESCRIPTION
    The script never pushes on its own. Publishing to nuget.org is done by
    .github/workflows/publish.yml when a v-tag is pushed, and that tag is set by
    release.ps1. This script stops at a verified package in the local feed.

    The shape check is the whole point. One of the reasons Avig.Ord exists is that DocX, NPOI
    and the commercial Word libraries pull in other packages. The check therefore fails on anything but
    a package with lib/net10.0/Avig.Ord.dll, a README and exactly two dependencies: Avig.Result and Avig.OpenXml (the package layer it shares with Avig.Kalkyl).

.PARAMETER NoBump
    Keep the current version instead of bumping the patch number. The workflow runs with this
    flag: the version is already set and tagged by release.ps1.

.PARAMETER SkipTests
    Skip the test run. Use when you have already run them.

.PARAMETER SkipVerify
    Skip the shape check of the nupkg. Do not do that without a specific reason.

.PARAMETER Feed
    Directory to copy the built package to for local smoke testing.
#>
param(
    [switch]$NoBump,
    [switch]$SkipTests,
    [switch]$SkipVerify,
    [string]$Feed = "G:\LocalNugetRepo"
)

$ErrorActionPreference = "Stop"

$root    = $PSScriptRoot
$project = Join-Path $root "src\Avig.Ord\Avig.Ord.csproj"
$tests   = Join-Path $root "tests\Avig.Ord.Tests\Avig.Ord.Tests.csproj"
$packDir = Join-Path $root "src\Avig.Ord\bin\Release"

Write-Host ""
Write-Host "=== Avig.Ord ===" -ForegroundColor Cyan
Write-Host ""

# --- [1/5] Version ---
Write-Host "[1/5] Version" -ForegroundColor Yellow

# XmlDocument.Load reads the file's own encoding declaration. Get-Content without -Encoding
# does not: in Windows PowerShell 5.1 UTF-8 is decoded as ANSI and Save writes it back as
# UTF-8, so every a-ring, a-umlaut, o-umlaut and dash doubles on every bump. That is how
# Medusa.Portal.csproj grew from 2 kB to 1.1 MB over seven version bumps.
$csproj = New-Object System.Xml.XmlDocument
$csproj.PreserveWhitespace = $true
$csproj.Load((Resolve-Path $project))

$currentVersion = [string](($csproj.Project.PropertyGroup | Where-Object { $_.Version }).Version)
if (-not $currentVersion) {
    Write-Host "      FEL: <Version> saknas i $project" -ForegroundColor Red
    exit 1
}

if ($NoBump) {
    $version = $currentVersion
    Write-Host "      $version (ingen bump)" -ForegroundColor Yellow
} else {
    $parts = $currentVersion.Split('.')
    $version = "$($parts[0]).$($parts[1]).$([int]$parts[2] + 1)"
    foreach ($pg in $csproj.Project.PropertyGroup) {
        if ($pg.Version) { $pg.Version = $version }
    }
    $csproj.Save((Resolve-Path $project))
    Write-Host "      $currentVersion -> $version" -ForegroundColor Yellow
}

# An already published package must never be replaced under the same version number: anyone
# who cached it keeps their copy and runs different code from anyone who fetches it fresh. Two
# different Avig.Ord in circulation without anything showing in the build. This applies
# after a bump too - if the bumped version is already in the feed, the csproj and the feed
# have drifted apart, and packing over it is exactly the wrong fix.
$existing = Join-Path $Feed "Avig.Ord.$version.nupkg"
if (Test-Path $existing) {
    Write-Host "      FEL: Avig.Ord $version finns redan i flodet." -ForegroundColor Red
    if ($NoBump) {
        Write-Host "           Hoj <Version> i stallet for att skriva over den." -ForegroundColor Red
    } else {
        Write-Host "           Bumpen landade pa en version som redan ar publicerad." -ForegroundColor Red
        Write-Host "           Satt <Version> i csproj:en till nasta lediga nummer." -ForegroundColor Red
    }
    exit 1
}

# --- [2/5] Tests ---
if ($SkipTests) {
    Write-Host "[2/5] Tester hoppas over" -ForegroundColor DarkYellow
} else {
    Write-Host "[2/5] Testar..." -ForegroundColor Yellow
    dotnet test $tests -c Release --nologo -v quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      Tester misslyckades - inget packat." -ForegroundColor Red
        exit 1
    }
    Write-Host "      Alla tester grona" -ForegroundColor Green
}

# --- [3/5] Pack ---
Write-Host "[3/5] Packar..." -ForegroundColor Yellow

# Old nupkgs in bin\Release must not be left behind and mistaken for the new one.
if (Test-Path $packDir) {
    Get-ChildItem -Path $packDir -Filter "*.nupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -Path $packDir -Filter "*.snupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
}

dotnet pack $project -c Release --nologo -v quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Pack misslyckades" -ForegroundColor Red
    exit 1
}

$package = Get-ChildItem -Path $packDir -Filter "Avig.Ord.$version.nupkg" | Select-Object -First 1
if (-not $package) {
    Write-Host "      Hittade inget paket for version $version" -ForegroundColor Red
    exit 1
}
Write-Host "      $($package.Name)" -ForegroundColor Green

# --- [4/5] Verify the package shape ---
if ($SkipVerify) {
    Write-Host "[4/5] Formkontroll hoppas over" -ForegroundColor DarkYellow
} else {
    Write-Host "[4/5] Verifierar paketets innehall..." -ForegroundColor Yellow

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($package.FullName)
    try {
        $entries = $zip.Entries | ForEach-Object { $_.FullName }
        $nuspecEntry = $zip.Entries | Where-Object { $_.FullName -eq "Avig.Ord.nuspec" }
        if (-not $nuspecEntry) {
            Write-Host "      FEL: nuspec saknas i paketet" -ForegroundColor Red
            exit 1
        }
        $reader = New-Object System.IO.StreamReader($nuspecEntry.Open())
        try { $nuspec = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $zip.Dispose()
    }

    if (-not ($entries | Where-Object { $_ -eq "lib/net10.0/Avig.Ord.dll" })) {
        Write-Host "      FEL: lib/net10.0/Avig.Ord.dll saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # The source generator for [OrdTemplate] must be included, otherwise no typed template
    # compiles for whoever uses the package.
    if (-not ($entries | Where-Object { $_ -eq "analyzers/dotnet/cs/Avig.Ord.Generators.dll" })) {
        Write-Host "      FEL: analyzers/dotnet/cs/Avig.Ord.Generators.dll saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # Avig.Result and Avig.OpenXml are the only allowed dependencies. Anything else - a zip, XML or image
    # library - is exactly what this package exists to avoid.
    $deps = @($nuspec.package.metadata.dependencies.group.dependency) | Where-Object { $_ -ne $null }
    $foreign = @($deps | Where-Object { $_.id -ne "Avig.Result" -and $_.id -ne "Avig.OpenXml" })
    if ($foreign.Count -gt 0) {
        Write-Host "      FEL: paketet drar in andra beroenden an Avig.Result och Avig.OpenXml:" -ForegroundColor Red
        $foreign | ForEach-Object { Write-Host "        $($_.id) $($_.version)" -ForegroundColor Red }
        exit 1
    }

    if (-not ($deps | Where-Object { $_.id -eq "Avig.Result" })) {
        Write-Host "      FEL: beroendet pa Avig.Result saknas i nuspec:en" -ForegroundColor Red
        exit 1
    }

    if (-not ($deps | Where-Object { $_.id -eq "Avig.OpenXml" })) {
        Write-Host "      FEL: beroendet pa Avig.OpenXml saknas i nuspec:en" -ForegroundColor Red
        exit 1
    }

    # The README is all that someone who finds the package in a feed has to go on.
    if (-not ($entries | Where-Object { $_ -eq "README.md" })) {
        Write-Host "      FEL: README.md saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # The license decides whether anyone may use the package. Without it, that is unclear.
    if (-not $nuspec.package.metadata.license) {
        Write-Host "      FEL: licens saknas i nuspec:en" -ForegroundColor Red
        exit 1
    }

    Write-Host "      lib/net10.0/Avig.Ord.dll, README.md och licens finns, bara Avig.Result och Avig.OpenXml som beroenden" -ForegroundColor Green
}

# --- [5/5] Copy to the local feed ---
Write-Host "[5/5] Kopierar till flodet..." -ForegroundColor Yellow
if (-not (Test-Path $Feed)) {
    New-Item -Path $Feed -ItemType Directory | Out-Null
}

Get-ChildItem -Path $packDir -Filter "Avig.Ord.$version.*nupkg" | ForEach-Object {
    Copy-Item $_.FullName -Destination $Feed -Force
    Write-Host "      $Feed\$($_.Name)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Avig.Ord v$version ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Paketet ligger i det lokala flodet. Till nuget.org gar det via en tagg:" -ForegroundColor Gray
Write-Host "    ./release.ps1 -Patch -Push" -ForegroundColor Gray
