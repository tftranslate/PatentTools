[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $root "build\PatentTools.dotm"
}

$baseDotm = Join-Path $root "template\PatentTools.base.dotm"
$vbaDir = Join-Path $root "source\vba"
$ribbonDir = Join-Path $root "source\ribbon\customUI"
$packageDir = Join-Path $root "source\package"
$buildDir = Split-Path -Parent $OutputPath
$outputPath = [System.IO.Path]::GetFullPath($OutputPath)

$requiredPaths = @(
    $baseDotm,
    $vbaDir,
    $ribbonDir,
    (Join-Path $packageDir "[Content_Types].xml"),
    (Join-Path $packageDir "_rels\.rels")
)

foreach ($path in $requiredPaths) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Erforderlicher Pfad fehlt: $path"
    }
}

New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
Copy-Item -LiteralPath $baseDotm -Destination $outputPath -Force

$word = $null
$document = $null

try {
    Write-Host "Starte Word und importiere VBA-Komponenten ..."

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($outputPath)

    $sourceFiles = @(
        Get-ChildItem -LiteralPath $vbaDir -Filter "*.bas" -File
        Get-ChildItem -LiteralPath $vbaDir -Filter "*.cls" -File
        Get-ChildItem -LiteralPath $vbaDir -Filter "*.frm" -File
    ) | Sort-Object Name

    if ($sourceFiles.Count -eq 0) {
        throw "Keine importierbaren VBA-Dateien in '$vbaDir' gefunden."
    }

    foreach ($file in $sourceFiles) {
        Write-Host "  Importiere: $($file.Name)"
        [void]$document.VBProject.VBComponents.Import($file.FullName)
    }

    $document.Save()
    $document.Close()
    $document = $null
}
finally {
    if ($null -ne $document) {
        try { $document.Close([ref]$false) } catch {}
    }

    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Word muss die Datei freigegeben haben, bevor das DOTM-Paket verändert wird.
Start-Sleep -Milliseconds 500

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = $null

try {
    Write-Host "Bette RibbonX-Ressourcen ein ..."

    $zip = [System.IO.Compression.ZipFile]::Open(
        $outputPath,
        [System.IO.Compression.ZipArchiveMode]::Update
    )

    function Remove-ZipEntry {
        param(
            [System.IO.Compression.ZipArchive]$Archive,
            [string]$EntryName
        )

        $entry = $Archive.GetEntry($EntryName)
        if ($null -ne $entry) {
            $entry.Delete()
        }
    }

    function Add-ZipFile {
        param(
            [System.IO.Compression.ZipArchive]$Archive,
            [string]$SourcePath,
            [string]$EntryName
        )

        Remove-ZipEntry -Archive $Archive -EntryName $EntryName

        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $Archive,
            $SourcePath,
            $EntryName,
            [System.IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null
    }

    # Vorhandenen Custom-UI-Part entfernen, falls die Basisvorlage doch einen enthält.
    $oldEntries = @(
        $zip.Entries |
        Where-Object {
            $_.FullName -eq "customUI/" -or
            $_.FullName -like "customUI/*"
        }
    )

    foreach ($entry in $oldEntries) {
        $entry.Delete()
    }

    # Vollständigen Ribbon-Quellordner unter customUI/ in das Paket kopieren.
    Get-ChildItem -LiteralPath $ribbonDir -Recurse -File | ForEach-Object {
        $relativePath = $_.FullName.Substring($ribbonDir.Length).TrimStart('\', '/')
        $entryName = "customUI/" + ($relativePath -replace '\\', '/')

        Write-Host "  Ribbon-Datei: $entryName"
        Add-ZipFile -Archive $zip -SourcePath $_.FullName -EntryName $entryName
    }

    # Die zwei gespeicherten Paket-Integrationsdateien einsetzen.
    Add-ZipFile -Archive $zip `
        -SourcePath (Join-Path $packageDir "[Content_Types].xml") `
        -EntryName "[Content_Types].xml"

    Add-ZipFile -Archive $zip `
        -SourcePath (Join-Path $packageDir "_rels\.rels") `
        -EntryName "_rels/.rels"
}
finally {
    if ($null -ne $zip) {
        $zip.Dispose()
    }
}

Write-Host ""
Write-Host "Build erfolgreich abgeschlossen:"
Write-Host "  $outputPath"