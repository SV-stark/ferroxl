<#
.SYNOPSIS
    Publishes ferroxl and ferroxl-mcp to crates.io, in the only order that works.

.DESCRIPTION
    The two crates have to be published one after the other. `ferroxl-mcp` depends on
    `ferroxl` by version, and crates.io will not accept a package whose dependency it
    cannot resolve - so `ferroxl` has to exist on the registry before `ferroxl-mcp` is
    packaged. This publishes the library, waits for the registry index to catch up, then
    publishes the server.

    Both versions come from Cargo.toml and must agree. The script refuses to run against
    a version that is already published, because crates.io does not allow replacing a
    published release and the resulting failure says so in a way that is easy to miss.

.PARAMETER DryRun
    Package and verify both crates without uploading. `ferroxl-mcp` still cannot be
    packaged before `ferroxl` is on the registry, so a dry run of the first release
    reports that as an expected failure rather than a problem.

.EXAMPLE
    ./scripts/publish.ps1 -DryRun
    ./scripts/publish.ps1
#>
[CmdletBinding()]
param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-Location (Join-Path $PSScriptRoot '..')

function Invoke-Native {
    # The argument list is passed as one array and splatted, rather than as trailing
    # `ValueFromRemainingArguments`. PowerShell tries to bind anything starting with a
    # dash against the function's own parameters first, so `cargo publish -p ferroxl`
    # became an ambiguity error about `-ProgressAction` before cargo ever ran.
    param([Parameter(Position = 0)][string]$Exe, [Parameter(Position = 1)][string[]]$NativeArgs)
    Write-Host "$Exe $($NativeArgs -join ' ')" -ForegroundColor DarkGray
    & $Exe @NativeArgs
    if ($LASTEXITCODE -ne 0) {
        throw "$Exe $($NativeArgs -join ' ') failed with exit code $LASTEXITCODE"
    }
}

# -- Refuse to publish something that is not what was tested --------------------------------

if (git status --porcelain) {
    throw "The working tree is not clean. Commit or stash first: crates.io publishes a commit, and publishing a dirty tree publishes something nobody reviewed."
}

$branch = git rev-parse --abbrev-ref HEAD
$tag = git tag --points-at HEAD | Where-Object { $_ -like 'v*' } | Select-Object -First 1
if (-not $tag) {
    Write-Warning "HEAD carries no v* tag. A release usually goes out with one, because the GitHub release workflow is triggered by it."
}
else {
    Write-Host "tagging $tag" -ForegroundColor DarkGray
}

$version = (Select-String -Path Cargo.toml -Pattern '^version = "(.+)"$').Matches[0].Groups[1].Value
Write-Host "publishing $version" -ForegroundColor Cyan

# -- Refuse to publish a version that already exists ----------------------------------------

# Checked through the crates.io API rather than `cargo search`, which reports the newest
# version rather than whether a specific one exists.
function Test-Published {
    param([string]$Crate, [string]$Version)
    try {
        $response = Invoke-RestMethod -Uri "https://crates.io/api/v1/crates/$Crate/$Version" -Headers @{ 'User-Agent' = 'ferroxl-publish' } -ErrorAction Stop
        return $null -ne $response.version
    }
    catch {
        # 404 is the answer we expect for a version that has not been published.
        return $false
    }
}

function Test-CargoCredential {
    <#
    .SYNOPSIS
        Whether cargo can authenticate against crates.io without being told a token.
    #>
    if ($env:CARGO_REGISTRY_TOKEN) {
        return $true
    }
    $home = if ($env:CARGO_HOME) { $env:CARGO_HOME } else { Join-Path $HOME '.cargo' }
    $file = Join-Path $home 'credentials.toml'
    if (-not (Test-Path $file)) {
        return $false
    }
    # The value is never read, only whether a token line exists.
    return [bool](Select-String -Path $file -Pattern '^\s*token\s*=' -Quiet)
}

# -- Publish ---------------------------------------------------------------------------------

$crates = @('ferroxl', 'ferroxl-mcp')

foreach ($crate in $crates) {
    if (Test-Published -Crate $crate -Version $version) {
        throw "$crate $version is already on crates.io. A published version cannot be replaced, so bump Cargo.toml instead of retrying."
    }
}

if ($DryRun) {
    Invoke-Native 'cargo' @('publish', '-p', $crates[0], '--dry-run', '--locked')
    Write-Host ""
    Write-Host "dry run: $version packaged and verified for $($crates[0])." -ForegroundColor Yellow
    Write-Host "$($crates[1]) cannot be dry-run until $($crates[0]) $version is on the registry - that check is the point of the sequencing, not a defect." -ForegroundColor Yellow
    exit 0
}

# Cargo reads a token from the environment or from `<CARGO_HOME>/credentials.toml`. Checking
# only the environment reported "no token" on a machine that had one on disk, which is worse
# than not checking: it sends the reader off to create a second token they do not need.
if (-not (Test-CargoCredential)) {
    throw "No crates.io credential found. Create a token at https://crates.io/settings/tokens with the 'publish:new' scope for the ferroxl-2 organisation, then either run 'gh secret set CARGO_REGISTRY_TOKEN --repo SV-stark/ferroxl' for CI, or set `$env:CARGO_REGISTRY_TOKEN for this session."
}

foreach ($crate in $crates) {
    Invoke-Native 'cargo' @('publish', '-p', $crate, '--locked')

    if ($crate -eq $crates[0]) {
        # The registry publishes the version before its index entry is readable, and
        # `ferroxl-mcp` cannot be packaged until that entry exists. Publishing straight on
        # fails with "no matching package named `ferroxl` found", which reads like a
        # version mistake rather than a race.
        Write-Host "waiting for the crates.io index to list $crate $version" -ForegroundColor DarkGray
        $deadline = (Get-Date).AddMinutes(10)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 15
            if (Test-Published -Crate $crate -Version $version) {
                Write-Host "$crate $version is indexed" -ForegroundColor Green
                break
            }
        }
        if ((Get-Date) -ge $deadline) {
            throw "$crate $version was uploaded but has not appeared in the index after 10 minutes. Check https://crates.io/crates/$crate before retrying; re-running this script will not help while the version already exists."
        }
    }
}

foreach ($crate in $crates) {
    if (-not (Test-Published -Crate $crate -Version $version)) {
        throw "$crate $version did not appear on crates.io. Something failed after the upload reported success."
    }
    Write-Host "https://crates.io/crates/$crate/$version" -ForegroundColor Green
}