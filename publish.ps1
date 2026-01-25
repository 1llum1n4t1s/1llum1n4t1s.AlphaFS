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
    
    $result = dotnet nuget push "$pkgPath" --api-key $apiKey --source https://api.nuget.org/v3/index.json 2>&1
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