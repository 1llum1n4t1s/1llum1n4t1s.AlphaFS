# NuGet へ公開する。
# API キーは -ApiKey 引数か、環境変数 NUGET_API_KEY で渡す。
param
(
    [string] $ApiKey = $env:NUGET_API_KEY
)

Write-Host $PSScriptRoot

if (-not $ApiKey)
{
    throw "Need to set the API key first (-ApiKey or `$env:NUGET_API_KEY)"
}

# 公開するバージョンは Directory.Build.props の <Version> を正本とする。
# mtime が最新の *.nupkg を拾う方式だと、artifacts/ に古い版や別 ID のパッケージが
# 残っていた場合にそれを公開してしまう (NuGet はバージョンを永久予約するため取り返しがつかない)。
# CI の "Verify Version matches branch name" と同じ読み取り方を使い、照合済みの版だけを push する。
$propsPath = Join-Path $PSScriptRoot "Directory.Build.props"
$version = ([xml](Get-Content $propsPath)).Project.PropertyGroup.Version

if (-not $version)
{
    Write-Error "Could not read <Version> from $propsPath"
    exit 1
}

$packageId = "1llum1n4t1s.AlphaFS"
$folder = Join-Path $PSScriptRoot "artifacts"
$pkg = "$packageId.$version.nupkg"
$pkgPath = Join-Path $folder $pkg

if (-not (Test-Path -LiteralPath $pkgPath))
{
    Write-Error "Package not found: $pkgPath (run 'dotnet pack' first)"
    exit 1
}

Write-Host "publishing $pkg"
Write-Host "Package path: $pkgPath"

# --skip-duplicate: 既に公開済みのバージョンを再 push しても 409 で CI を落とさず exit 0 で抜ける。
# これにより誤再 push と本物の公開失敗 (API key 失効 / NuGet downtime 等) を CI ログで区別しやすくなる。
$result = dotnet nuget push "$pkgPath" --api-key $ApiKey --source https://api.nuget.org/v3/index.json --skip-duplicate 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0)
{
    Write-Host "Error output:"
    Write-Host $result
    Write-Error "Failed to publish $pkg (exit code: $exitCode)"
    exit $exitCode
}

Write-Host "Successfully published $pkg"
