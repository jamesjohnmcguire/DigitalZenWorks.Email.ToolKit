# Upload executable releases or publish library packages through GitHub Actions.
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][ValidateSet('release', 'publish')][string]$Action,
    [Parameter(Mandatory = $true)][string]$Repository,
    [Parameter(Mandatory = $true)][string]$Version,
    [Parameter(Mandatory = $true)][string]$Commit,
    [string]$Package,
    [string]$Archive,
    [string]$CommitRecord,
    [Parameter(Mandatory = $true)][string]$Notes
)

$ErrorActionPreference = 'Stop'

try
{
    if ($Action -eq 'release')
    {
        $releaseAssets = @($Archive)
    }
    else
    {
        $releaseAssets = @($Package, $CommitRecord)
        $stream = [System.IO.File]::OpenRead($Package)
        $algorithm = [System.Security.Cryptography.SHA256]::Create()
        try
        {
            $hash = $algorithm.ComputeHash($stream)
            $digest = [System.BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()
        }
        finally
        {
            $stream.Dispose()
            $algorithm.Dispose()
        }
    }
    foreach ($asset in $releaseAssets)
    {
        if ([string]::IsNullOrEmpty($asset) -or [System.IO.File]::Exists($asset) -eq $false)
        {
            throw "Release artifact is missing: $asset"
        }
    }
    # Reuse an existing release on retry; never overwrite published assets.
    $tag = "v$Version"
    $savedErrorActionPreference = $ErrorActionPreference
    try
    {
        # Windows PowerShell treats redirected native stderr as an error.
        $ErrorActionPreference = 'Continue'
        $releaseJson = & gh release view $tag --repo $Repository --json isDraft,targetCommitish,assets 2>$null
        $releaseExitCode = $LASTEXITCODE
    }
    finally
    {
        $ErrorActionPreference = $savedErrorActionPreference
    }
    if ($releaseExitCode -eq 0)
    {
        $release = $releaseJson | ConvertFrom-Json
        if ($release.isDraft)
        {
            if ($release.targetCommitish -ne $Commit)
            {
                throw 'The existing draft targets another commit. Review its target before retrying.'
            }
            & gh release upload $tag @releaseAssets --repo $Repository --clobber
            if ($LASTEXITCODE -ne 0)
            {
                throw 'Unable to upload the reviewed draft release artifacts.'
            }
        }
        else
        {
            if ($Action -eq 'release')
            {
                throw 'This GitHub release is already published. Its executable will not be overwritten.'
            }
            $taggedCommit = & gh api "repos/$Repository/commits/$tag" --jq .sha
            if ($LASTEXITCODE -ne 0 -or $taggedCommit -ne $Commit)
            {
                throw 'The published release tag does not match the current source commit.'
            }
            $existingNames = @($release.assets | ForEach-Object { $_.name })
            $presentAssets = @($releaseAssets | Where-Object { $existingNames -contains [System.IO.Path]::GetFileName($_) })
            if ($presentAssets.Count -eq 0)
            {
                & gh release upload $tag @releaseAssets --repo $Repository
                if ($LASTEXITCODE -ne 0)
                {
                    throw 'Unable to upload the package assets to the existing GitHub release.'
                }
            }
            elseif ($presentAssets.Count -ne $releaseAssets.Count)
            {
                throw 'Package assets are incomplete on the published release. Review them before retrying.'
            }
            else
            {
                Write-Host 'Using existing package assets; the workflow will verify their contents.'
            }
        }
    }
    else
    {
        & gh release create $tag @releaseAssets --draft --target $Commit --notes-file $Notes --repo $Repository
        if ($LASTEXITCODE -ne 0)
        {
            throw 'Unable to create the draft release and upload its artifacts.'
        }
    }

    if ($Action -eq 'release')
    {
        Write-Host 'Executable uploaded to the GitHub draft release. No NuGet publication was requested.'
        exit 0
    }

    $requestId = [guid]::NewGuid().ToString('N')
    $title = "NuGet $Version / $requestId"
    Write-Host "Requesting publication of $Version from $Commit."
    Write-Host "Package SHA-256: $digest"
    & gh workflow run build.yml --ref main --repo $Repository -F publish=true -f "version=$Version" -f "commit=$Commit" -f "sha256=$digest" -f "request_id=$requestId"
    if ($LASTEXITCODE -ne 0)
    {
        throw 'Unable to dispatch the publishing workflow.'
    }

    $runId = $null
    for ($attempt = 0; $attempt -lt 12; $attempt++)
    {
        $runsJson = & gh run list --repo $Repository --workflow build.yml --event workflow_dispatch --branch main --limit 50 --json databaseId,displayTitle
        if ($LASTEXITCODE -ne 0)
        {
            throw "Workflow requested, but its status could not be read. Find run '$title' in GitHub Actions before retrying."
        }
        $runs = $runsJson | ConvertFrom-Json
        $run = $runs | Where-Object { $_.displayTitle -eq $title } | Select-Object -First 1
        if ($null -ne $run)
        {
            $runId = $run.databaseId
            break
        }
        Start-Sleep -Seconds 5
    }
    if ($null -eq $runId)
    {
        throw "Workflow requested, but not visible yet. Find run '$title' in GitHub Actions before retrying."
    }

    Write-Host "Following workflow run $runId."
    & gh run watch $runId --repo $Repository --exit-status
    if ($LASTEXITCODE -ne 0)
    {
        throw "Publication did not succeed. Inspect: gh run view $runId --repo $Repository --log-failed"
    }
}
catch
{
    Write-Host $_.Exception.Message
    exit 1
}
exit 0
