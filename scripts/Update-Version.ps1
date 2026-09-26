[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$NewVersion,
    
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

Write-Host "`n=== Updating PatentTools Version ===" -ForegroundColor Cyan
Write-Host "Current version: $(Get-Content (Join-Path $root 'VERSION') | Select-Object -First 1)" -ForegroundColor Gray
Write-Host "New version: $NewVersion" -ForegroundColor Yellow

# Files to update with the new version
$filesToUpdate = @(
    @{ Path = Join-Path $root 'README.md'; Pattern = 'v \d+\.\d+\.\d+'; Replacement = "v $NewVersion" },
    @{ Path = Join-Path $root 'CHANGELOG.md'; Pattern = '## \[\d+\.\d+\.\d+\]'; Replacement = "## [$NewVersion]" }
)

# Update VERSION file (main source of truth)
$versionFile = Join-Path $root 'VERSION'
if (-not $WhatIf) {
    Set-Content -Path $versionFile -Value $NewVersion -NoNewline
    Write-Host "✅ Updated: VERSION" -ForegroundColor Green
} else {
    Write-Host "📝 WhatIf: Would update: VERSION to $NewVersion" -ForegroundColor Yellow
}

# Update other files
foreach ($fileInfo in $filesToUpdate) {
    if (Test-Path $fileInfo.Path) {
        $content = Get-Content $fileInfo.Path -Raw
        $newContent = $content -replace $fileInfo.Pattern, $fileInfo.Replacement
        
        if ($newContent -ne $content) {
            if (-not $WhatIf) {
                Set-Content -Path $fileInfo.Path -Value $newContent -NoNewline
                Write-Host "✅ Updated: $($fileInfo.Path)" -ForegroundColor Green
            } else {
                Write-Host "📝 WhatIf: Would update: $($fileInfo.Path)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "⚠️  No changes needed: $($fileInfo.Path)" -ForegroundColor Gray
        }
    } else {
        Write-Host "⚠️  File not found: $($fileInfo.Path)" -ForegroundColor Yellow
    }
}

Write-Host "`n=== Version update complete ===" -ForegroundColor Cyan
if ($WhatIf) {
    Write-Host "Run without -WhatIf to apply changes" -ForegroundColor Gray
}
