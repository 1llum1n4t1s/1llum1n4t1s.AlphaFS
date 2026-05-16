# publishes to NuGet
# $apiKey needs to be already set up with NuGet publishing key
Write-Host $PSScriptRoot

if (-not $apiKey)
{
    throw "Need to set the API key first"
}

# publish AlphaFS package
$folder = "$PSScriptRoot\artifacts"
$recent = Get-ChildItem -Path $folder -Filter "*.nupkg" -Recurse | Sort-Object LastWriteTime | Select-Object -Last 1
if ($recent)
{
    $pkg = $recent.Name
    $pkgPath = $recent.FullName
    Write-Host "publishing $pkg"
    Write-Host "Package path: $pkgPath"
    Write-Host "API key length: $($apiKey.Length)"
    
    # --skip-duplicate: 既に公開済みのバージョンを再 push しても 409 で CI を落とさず exit 0 で抜ける。
    # これにより誤再 push と本物の公開失敗 (API key 失効 / NuGet downtime 等) を CI ログで区別しやすくなる。
    $result = dotnet nuget push "$pkgPath" --api-key $apiKey --source https://api.nuget.org/v3/index.json --skip-duplicate 2>&1
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -ne 0)
    {
        Write-Host "Error output:"
        Write-Host $result
        Write-Error "Failed to publish $pkg (exit code: $exitCode)"
        exit $exitCode
    }
    else
    {
        Write-Host "Successfully published $pkg"
    }
}
else
{
    Write-Error "No package found in $folder"
    exit 1
}