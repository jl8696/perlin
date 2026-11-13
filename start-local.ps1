$ErrorActionPreference = "Stop"
$env:NODE_OPTIONS = "--use-system-ca"

if (-not (Test-Path -LiteralPath ".\node_modules")) {
  npm install
  if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed."
  }
}

npm start
