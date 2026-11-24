#Requires -Version 7.0
<#
.SYNOPSIS
    Makes a release: bumps the version, verifies the package, commits and tags.

.DESCRIPTION
    The tag is what publishes. A pushed vX.Y.Z runs .github/workflows/publish.yml,
    which builds the package again on a runner and pushes it to nuget.org via Trusted
    Publishing.

    The script does everything up to that point and stops, because a tag that has not left
    the machine can still be deleted, while a published version can never be replaced. With
    -Push it goes all the way in one command.

    The package is packed and shape-checked BEFORE the tag is created, so a release that
    cannot be built never becomes a tag. If the check fails, the working tree is restored.

    Versions carry at ten instead of counting on: 1.9.0 becomes 1.10.0, but 1.10.0
    becomes 2.0.0, because ten is the last value a segment takes and not the first it
    refuses. Patch carries into minor the same way, which can carry on into major: 1.10.10
    becomes 2.0.0.

.PARAMETER Patch
    Bump the third number. The default when no bump is named - a release of this package is
    usually a fix, not new surface. The type sits at the bottom of other people's call chains,
    and new public surface there is rarely what you want.

.PARAMETER Minor
    Bump the second number and reset the third. For new members on Result.

.PARAMETER Major
    Bump the first number and reset the rest. For a change that breaks existing code.

.PARAMETER Version
    Set an explicit version instead of bumping, e.g. 2.0.0 or 2.0.0-rc.1.

.PARAMETER Push
    Push the branch and the tag once everything has passed, which starts the publishing.

.PARAMETER DryRun
    Report what would happen and change nothing.

.PARAMETER Branch
    The branch a release may be made from. Default main.

.EXAMPLE
    ./release.ps1
    1.0.0 becomes 1.0.1, verified, committed and tagged locally.

.EXAMPLE
    ./release.ps1 -Minor -Push
    New members went in: 1.0.1 becomes 1.1.0, and publishing starts.

.EXAMPLE
    ./release.ps1 -DryRun
    Tells which version a normal run would produce, and changes nothing.
#>
param(
    [switch]$Patch,
    [switch]$Minor,
    [switch]$Major,
    [string]$Version,
    [switch]$Push,
    [switch]$DryRun,
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"

# Ten is the last value a version segment takes, not the first it refuses. Change this
# single number to change the whole scheme.
$SegmentCeiling = 10

$root    = $PSScriptRoot
$project = Join-Path $root "src\Avig.Ord\Avig.Ord.csproj"
$publish = Join-Path $root "publish.ps1"

function Fail([string]$Message) {
    Write-Host "      $Message" -ForegroundColor Red
    exit 1
}

function Step-Version([int]$MajorPart, [int]$MinorPart, [int]$PatchPart, [string]$Part) {
    switch ($Part) {
        "major" { $MajorPart++; $MinorPart = 0; $PatchPart = 0 }
        "minor" { $MinorPart++; $PatchPart = 0 }
        default { $PatchPart++ }
    }

    # Carried in order, so a patch bump can travel all the way: 1.10.10 becomes 2.0.0.
    if ($PatchPart -gt $SegmentCeiling) {
        $PatchPart = 0
        $MinorPart++
    }

    if ($MinorPart -gt $SegmentCeiling) {
        $MinorPart = 0
        $MajorPart++
    }

    return "$MajorPart.$MinorPart.$PatchPart"
}

function Save-Text([string]$Path, [string]$Text) {
    # UTF-8 without BOM, and the line endings already in the string. .gitattributes locks this
    # repo to CRLF, so nothing here may normalize to LF.
    [System.IO.File]::WriteAllText($Path, $Text, [System.Text.UTF8Encoding]::new($false))
}

Write-Host ""
Write-Host "=== Avig.Ord release ===" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "    DRY RUN - ingenting skrivs, commitas eller pushas" -ForegroundColor Yellow
}

Write-Host ""

# --- [1/5] The repo must be in a state a tag can describe ---
Write-Host "[1/5] Kontrollerar arbetskatalogen..." -ForegroundColor Yellow

$currentBranch = (git -C $root rev-parse --abbrev-ref HEAD).Trim()
if ($currentBranch -ne $Branch) {
    Fail "Pa grenen '$currentBranch', inte '$Branch'. Slapp fran $Branch, eller ange -Branch."
}

if ((git -C $root status --porcelain)) {
    Fail "Arbetskatalogen har ocommitade andringar. En tagg maste beskriva ett commitat lage."
}

git -C $root fetch origin --tags --quiet
if ($LASTEXITCODE -ne 0) {
    Fail "Nadde inte origin."
}

$behind = [int](git -C $root rev-list --count "$Branch..origin/$Branch")
if ($behind -gt 0) {
    Fail "$Branch ligger $behind commit(s) efter origin/$Branch. Pulla forst."
}

Write-Host "      ren, pa $Branch, i fas med origin" -ForegroundColor Green

# --- [2/5] Work out the new version ---
Write-Host "[2/5] Version" -ForegroundColor Yellow

$csproj = Get-Content $project -Raw
$currentMatch = ([regex]'<Version>([^<]+)</Version>').Match($csproj)
if (-not $currentMatch.Success) {
    Fail "Hittade ingen <Version> i csproj:en."
}

$current = $currentMatch.Groups[1].Value

$named = @($Patch, $Minor, $Major, [bool]$Version) | Where-Object { $_ }
if ($named.Count -gt 1) {
    Fail "Valj en av -Patch, -Minor, -Major eller -Version."
}

if ($Version) {
    if ($Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.\-]+)?$') {
        Fail "'$Version' ar ingen version. Forvantade 1.2.3 eller 1.2.3-rc.1."
    }

    $next = $Version
    $part = "explicit"
} else {
    $core = ($current -split '-')[0]
    $parts = $core.Split('.')
    if ($parts.Count -ne 3) {
        Fail "Kan inte hoja '$current' - forvantade tre tal."
    }

    # Patch by default: a release of this package is usually a fix.
    $part = if ($Major) { "major" } elseif ($Minor) { "minor" } else { "patch" }
    $next = Step-Version ([int]$parts[0]) ([int]$parts[1]) ([int]$parts[2]) $part
}

if ($next -eq $current) {
    Fail "$next ar redan den nuvarande versionen."
}

$currentCore = [version]($current -split '-')[0]
$nextCore = [version]($next -split '-')[0]
if ($nextCore -lt $currentCore) {
    Fail "$next ar lagre an nuvarande $current. Versioner pa nuget.org gar bara uppat."
}

$tag = "v$next"

if ((git -C $root tag -l $tag)) {
    Fail "Taggen $tag finns redan lokalt. Ta bort den eller valj en annan version."
}

if ((git -C $root ls-remote --tags origin "refs/tags/$tag")) {
    Fail "Taggen $tag finns redan pa origin. En publicerad version gar inte att byta ut."
}

$carried = ""
if ($part -ne "explicit" -and $part -ne "major" -and $nextCore.Major -gt $currentCore.Major) {
    $carried = "   bar over i en ny major"
} elseif ($part -eq "patch" -and $nextCore.Minor -gt $currentCore.Minor) {
    $carried = "   bar over i en ny minor"
}

Write-Host "      $current -> $next   (tagg $tag)$carried" -ForegroundColor Green

# --- [3/5] Prove the release builds before it becomes a tag ---
Write-Host "[3/5] Verifierar paketet..." -ForegroundColor Yellow

if ($DryRun) {
    Write-Host "      hoppas over i en dry run" -ForegroundColor Yellow
} else {
    $updatedCsproj = ([regex]'<Version>[^<]*</Version>').Replace($csproj, "<Version>$next</Version>", 1)
    Save-Text $project $updatedCsproj

    # No pipeline wrapping: publish.ps1 reports via Write-Host, so the output goes straight
    # to the host and nothing would arrive here to indent.
    & $publish -NoBump
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      publish.ps1 misslyckades - aterstaller arbetskatalogen, ingen tagg skapad." -ForegroundColor Red
        git -C $root checkout -- $project
        exit 1
    }

    Write-Host "      tester, pack och formkontroll gick igenom" -ForegroundColor Green
}

# --- [4/5] Commit and tag ---
Write-Host "[4/5] Commit och tagg" -ForegroundColor Yellow

$message = "Release $tag"

if ($DryRun) {
    Write-Host "      skulle commita '$message' och tagga $tag" -ForegroundColor Yellow
} else {
    git -C $root add -- $project
    git -C $root commit --quiet -m $message
    if ($LASTEXITCODE -ne 0) {
        Fail "Commit misslyckades."
    }

    git -C $root tag -a $tag -m "Avig.Ord $next"
    if ($LASTEXITCODE -ne 0) {
        Fail "Taggningen misslyckades. Release-commiten ligger kvar; ta bort den eller tagga om for hand."
    }

    Write-Host "      commitad och taggad $tag" -ForegroundColor Green
}

# --- [5/5] The push is what publishes ---
Write-Host "[5/5] Push" -ForegroundColor Yellow

if ($DryRun) {
    Write-Host "      skulle pusha $Branch och $tag" -ForegroundColor Yellow
} elseif ($Push) {
    git -C $root push origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Fail "Push av $Branch misslyckades. Taggen ar fortfarande lokal; ingenting publicerat."
    }

    git -C $root push origin $tag
    if ($LASTEXITCODE -ne 0) {
        Fail "Push av $tag misslyckades. Ingenting publicerat."
    }

    Write-Host "      pushad - publiceringen startar" -ForegroundColor Green
} else {
    Write-Host "      inte pushad (ange -Push for att gora det harifran)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== $tag forberedd ===" -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "Ingenting andrades. Kor om utan -DryRun." -ForegroundColor Cyan
} elseif ($Push) {
    Write-Host "Folj publiceringen:" -ForegroundColor Cyan
    Write-Host "    https://github.com/marcuslindh/Avig.Ord/actions" -ForegroundColor Gray
} else {
    Write-Host "Ingenting har lamnat maskinen an. For att publicera:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    git push origin $Branch && git push origin $tag" -ForegroundColor Gray
    Write-Host ""
    Write-Host "For att angra i stallet:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    git tag -d $tag; git reset --hard HEAD~1" -ForegroundColor Gray
}

Write-Host ""
