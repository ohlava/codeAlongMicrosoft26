<#
.SYNOPSIS
    Zkopíruje knihovny dokumentů ze vzorového webu na cílový, včetně složek
    a souborů.

.DESCRIPTION
    Odkazy na stránkách vzorového webu míří na soubory v jeho knihovnách.
    Bez těch souborů jsou v cílovém webu mrtvé - stránka je, ale nic se
    neotevře. Tenhle krok je doplní.

    Vychází z funkce Copy-PnPDocLibs ve script.ps1 verze 2.7 (autor Sergiu
    Nica). Proti ní tady navíc je režim náhledu, přeskakování souborů, které
    už v cíli jsou, a strop na velikost souboru.

    VÝCHOZÍ REŽIM JE DRY-RUN. Nic se nemaže ani nepřepisuje - soubor, který
    v cíli existuje, se přeskočí.

.PARAMETER SourceSiteUrl
    Vzorový web. Jen se z něj čte.

.PARAMETER TargetSiteUrl
    Web, na který se kopíruje.

.PARAMETER ClientId
    Client ID Entra ID aplikace. Když se nezadá, vezme se z PNP_CLIENT_ID.

.PARAMETER Libraries
    Které knihovny kopírovat. Výchozí je všechny kromě systémových.

.PARAMETER MaxFileSizeMB
    Soubory větší než tato hodnota se přeskočí a vypíšou. Výchozí 250 MB.

.PARAMETER Overwrite
    Přepíše i soubory, které v cíli už jsou. Bez toho se přeskočí.

.PARAMETER Apply
    Provede kopírování.

.EXAMPLE
    ./src/Copy-DocumentLibraries.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/P01" -TargetSiteUrl "https://contoso.sharepoint.com/sites/P42"

.EXAMPLE
    ./src/Copy-DocumentLibraries.ps1 -SourceSiteUrl "..." -TargetSiteUrl "..." -Apply

.NOTES
    Postup: docs/12-nastaveni-a-spusteni.md

    Soubory se stahují přes lokální disk, takže velká knihovna trvá dlouho.
    Náhled ukáže počet souborů i objem dat, než se do toho pustíte.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SourceSiteUrl,
    [Parameter(Mandatory = $true)] [string] $TargetSiteUrl,

    [string] $ClientId = $env:PNP_CLIENT_ID,
    [string] $OutputFolder = "./export",

    [string[]] $Libraries = @(),
    [int] $MaxFileSizeMB = 250,

    [switch] $Overwrite,
    [switch] $Apply
)

$ErrorActionPreference = "Stop"

$script:Warnings = @()

# Knihovny, které patří SharePointu, ne projektu.
$SystemLibraries = @(
    "Site Assets", "Websiteobjekte", "Prostředky webu",
    "Style Library", "Knihovna stylů",
    "Form Templates", "Formulářové šablony",
    "Site Pages", "Websiteseiten", "Stránky webu",
    "Documents", "Dokumente", "Dokumenty", "Shared Documents"
)

function Write-Step($Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Add-FileWarning($Message) {
    $script:Warnings += $Message
    Write-Warning $Message
}

function Get-SiteRelativeUrl($ServerRelativeUrl, $SiteRoot) {
    $relative = $ServerRelativeUrl
    if ($SiteRoot -and $SiteRoot -ne "/" -and $relative.StartsWith($SiteRoot)) {
        $relative = $relative.Substring($SiteRoot.Length)
    }
    return $relative.TrimStart("/")
}

function Format-Size($Bytes) {
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "{0:N0} kB" -f ($Bytes / 1KB)
}

# Všechny položky knihovny jedním dotazem, rozdělené na složky a soubory.
function Get-LibraryContent($List, $Connection) {
    $items = Get-PnPListItem -List $List.Id -PageSize 2000 -Connection $Connection `
        -Fields "FileRef", "FileLeafRef", "FSObjType", "File_x0020_Size"

    $folders = @()
    $files = @()

    foreach ($item in $items) {
        $values = $item.FieldValues
        if (-not $values.FileRef) { continue }

        if ($values.FSObjType -eq 1) {
            $folders += $values.FileRef
        }
        else {
            $size = 0
            if ($values.File_x0020_Size) { $size = [int64]$values.File_x0020_Size }

            $files += [pscustomobject]@{
                FileRef  = $values.FileRef
                FileName = $values.FileLeafRef
                Size     = $size
            }
        }
    }

    return [pscustomobject]@{ Folders = @($folders); Files = @($files) }
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "Modul PnP.PowerShell není nainstalovaný. Spusťte: Install-Module PnP.PowerShell -Scope CurrentUser"
}
if (-not $ClientId) {
    throw "Chybí ClientId. Zadejte -ClientId nebo nastavte PNP_CLIENT_ID."
}
if ($SourceSiteUrl.TrimEnd("/") -eq $TargetSiteUrl.TrimEnd("/")) {
    throw "Zdrojový a cílový web jsou tentýž."
}

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
$workFolder = Join-Path $OutputFolder "files-temp"
if (Test-Path $workFolder) { Remove-Item -Path $workFolder -Recurse -Force }
New-Item -Path $workFolder -ItemType Directory -Force | Out-Null

Write-Step "Připojuji se ke vzorovému webu"
Write-Host "  $SourceSiteUrl"
$sourceConnection = Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

Write-Step "Připojuji se k cílovému webu"
Write-Host "  $TargetSiteUrl"
$targetConnection = Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

$sourceWeb = Get-PnPWeb -Connection $sourceConnection
$targetWeb = Get-PnPWeb -Connection $targetConnection

# ============================================================
# Které knihovny
# ============================================================

Write-Step "Hledám knihovny ve vzorovém webu"

$sourceLibraries = @(Get-PnPList -Includes RootFolder -Connection $sourceConnection | Where-Object {
    $_.BaseTemplate -eq 101 -and -not $_.Hidden
})

if ($Libraries.Count -gt 0) {
    $sourceLibraries = @($sourceLibraries | Where-Object { $Libraries -contains $_.Title })
}
else {
    $sourceLibraries = @($sourceLibraries | Where-Object { $SystemLibraries -notcontains $_.Title })
}

if ($sourceLibraries.Count -eq 0) {
    Write-Host "  žádné knihovny ke kopírování"
    Write-Host ""
    Write-Host "  Výchozí filtr vynechává systémové knihovny i výchozí 'Dokumenty' -" -ForegroundColor DarkGray
    Write-Host "  ta se plní ze struktury v Excelu. Konkrétní knihovnu si vyžádejte" -ForegroundColor DarkGray
    Write-Host "  parametrem -Libraries \"Dokumenty\"." -ForegroundColor DarkGray
    return
}

foreach ($library in $sourceLibraries) {
    Write-Host "  - $($library.Title)"
}

# ============================================================
# Plán
# ============================================================

Write-Step "Zjišťuji obsah"

$plan = foreach ($library in $sourceLibraries) {
    $content = Get-LibraryContent $library $sourceConnection
    $totalSize = ($content.Files | Measure-Object -Property Size -Sum).Sum
    if (-not $totalSize) { $totalSize = 0 }

    Write-Host ("  {0,-30} {1,4} složek, {2,5} souborů, {3}" -f `
        $library.Title, $content.Folders.Count, $content.Files.Count, (Format-Size $totalSize))

    [pscustomobject]@{
        Library   = $library
        Content   = $content
        TotalSize = $totalSize
    }
}

$grandTotal = ($plan | Measure-Object -Property TotalSize -Sum).Sum
if (-not $grandTotal) { $grandTotal = 0 }
$fileCount = ($plan | ForEach-Object { $_.Content.Files.Count } | Measure-Object -Sum).Sum

Write-Host ""
Write-Host "  celkem: $fileCount souborů, $(Format-Size $grandTotal)"

if (-not $Apply) {
    Write-Step "DRY-RUN - nic se nekopírovalo"
    Write-Host "Kopírování jde přes lokální disk, takže počítejte s časem úměrným objemu." -ForegroundColor Yellow
    Write-Host "Až bude výpis v pořádku, spusťte stejný příkaz s přepínačem -Apply." -ForegroundColor Yellow
    return
}

# ============================================================
# Kopírování
# ============================================================

$maxBytes = $MaxFileSizeMB * 1MB
$copied = 0
$skipped = 0
$failed = 0

foreach ($entry in $plan) {
    $library = $entry.Library
    Write-Step "Knihovna: $($library.Title)"

    $sourceRoot = $library.RootFolder.ServerRelativeUrl
    $librarySiteRelative = Get-SiteRelativeUrl $sourceRoot $sourceWeb.ServerRelativeUrl

    # Cílová knihovna musí existovat. Zakládá se, pokud chybí - to je přidání,
    # ne přepis, takže je to v pořádku i při opakovaném běhu.
    $targetLibrary = Get-PnPList -Identity $library.Title -Connection $targetConnection -ErrorAction SilentlyContinue
    if (-not $targetLibrary) {
        $targetLibrary = New-PnPList -Title $library.Title -Template DocumentLibrary -Connection $targetConnection
        Write-Host "  knihovna vytvořena" -ForegroundColor Green
    }

    # Složky nejdřív, od nejkratší cesty, ať existují dřív než soubory v nich.
    foreach ($folderRef in ($entry.Content.Folders | Sort-Object Length)) {
        $relative = Get-SiteRelativeUrl $folderRef $sourceWeb.ServerRelativeUrl
        try {
            Resolve-PnPFolder -SiteRelativePath $relative -Connection $targetConnection -ErrorAction Stop | Out-Null
        }
        catch {
            Add-FileWarning "Složku '$relative' nelze vytvořit: $($_.Exception.Message)"
        }
    }

    $existingFiles = @{}
    if (-not $Overwrite) {
        $targetContent = Get-LibraryContent $targetLibrary $targetConnection
        foreach ($file in $targetContent.Files) {
            $key = Get-SiteRelativeUrl $file.FileRef $targetWeb.ServerRelativeUrl
            $existingFiles[$key] = $true
        }
    }

    $index = 0
    foreach ($file in $entry.Content.Files) {
        $index++
        $relative = Get-SiteRelativeUrl $file.FileRef $sourceWeb.ServerRelativeUrl

        if (-not $Overwrite -and $existingFiles.ContainsKey($relative)) {
            $skipped++
            continue
        }

        if ($file.Size -gt $maxBytes) {
            Add-FileWarning "Soubor '$relative' má $(Format-Size $file.Size), víc než limit $MaxFileSizeMB MB. Přeskakuji."
            $skipped++
            continue
        }

        $targetFolder = $relative.Substring(0, [Math]::Max(0, $relative.Length - $file.FileName.Length - 1))
        if (-not $targetFolder) { $targetFolder = $librarySiteRelative }

        Write-Progress -Activity "Kopíruji $($library.Title)" `
            -Status "$index / $($entry.Content.Files.Count): $($file.FileName)" `
            -PercentComplete (($index / [Math]::Max(1, $entry.Content.Files.Count)) * 100)

        $localPath = Join-Path $workFolder $file.FileName

        try {
            Get-PnPFile -Url $file.FileRef -Path $workFolder -FileName $file.FileName `
                -AsFile -Force -Connection $sourceConnection -ErrorAction Stop | Out-Null

            Add-PnPFile -Path $localPath -Folder $targetFolder -NewFileName $file.FileName `
                -Connection $targetConnection -ErrorAction Stop | Out-Null

            $copied++
            Write-Host "  + $relative" -ForegroundColor Green
        }
        catch {
            $failed++
            Add-FileWarning "Soubor '$relative' nelze zkopírovat: $($_.Exception.Message)"
        }
        finally {
            if (Test-Path $localPath) { Remove-Item -Path $localPath -Force -ErrorAction SilentlyContinue }
        }
    }

    Write-Progress -Activity "Kopíruji $($library.Title)" -Completed
}

if (Test-Path $workFolder) { Remove-Item -Path $workFolder -Recurse -Force -ErrorAction SilentlyContinue }

Write-Step "Hotovo"
Write-Host "  zkopírováno: $copied, přeskočeno: $skipped, chyb: $failed"

if ($script:Warnings.Count -gt 0) {
    $script:Warnings | Set-Content -Path "$OutputFolder/files-warnings.txt" -Encoding UTF8
    Write-Host "  $($script:Warnings.Count) varování v $OutputFolder/files-warnings.txt" -ForegroundColor Yellow
}
