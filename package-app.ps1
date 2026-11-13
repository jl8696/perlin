$ErrorActionPreference = "Stop"
$env:NODE_OPTIONS = "--use-system-ca"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

if ([string]::IsNullOrWhiteSpace($env:ADDIN_PROD_URL)) {
  throw "Set ADDIN_PROD_URL to the final HTTPS directory before packaging."
}

npm run build
if ($LASTEXITCODE -ne 0) {
  throw "The production build failed."
}

npx office-addin-manifest validate .\dist\manifest.json
if ($LASTEXITCODE -ne 0) {
  throw "The production manifest failed validation."
}

$zipPath = Join-Path $projectRoot "dist\particle-background-app-package.zip"
if (Test-Path -LiteralPath $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::CreateNew)
$archive = $null

try {
  $archive = [System.IO.Compression.ZipArchive]::new(
    $stream,
    [System.IO.Compression.ZipArchiveMode]::Create,
    $false
  )

  $files = @(
    @{ Source = "dist\manifest.json"; Entry = "manifest.json" },
    @{ Source = "dist\assets\icon-32.png"; Entry = "assets/icon-32.png" },
    @{ Source = "dist\assets\icon-192.png"; Entry = "assets/icon-192.png" }
  )

  foreach ($file in $files) {
    $source = Join-Path $projectRoot $file.Source
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
      $archive,
      $source,
      $file.Entry,
      [System.IO.Compression.CompressionLevel]::Optimal
    ) | Out-Null
  }
}
finally {
  if ($archive) {
    $archive.Dispose()
  }
  $stream.Dispose()
}

Write-Output "Created $zipPath"
