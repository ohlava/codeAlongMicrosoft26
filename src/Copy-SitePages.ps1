<#
.SYNOPSIS
    Přenese ze vzorového webu stránky s webparty, obrázky na nich, domovskou
    stránku, vzhled webu a regionální nastavení.

.DESCRIPTION
    Doplňuje to, co ostatní skripty neumí. Vychází z funkcí GetAllPages,
    CopyWebparts2, AddImages, Copy-Navigation a Copy-Regionalsettings
    v původním script.ps1 a používá stejný osvědčený postup:

        Export-PnPPage  ->  Invoke-PnPSiteTemplate

    Proti původnímu skriptu tady navíc je režim náhledu, jedna chyba nezastaví
    zbytek a nic se nemaže.

    VÝCHOZÍ REŽIM JE DRY-RUN. Bez přepínače -Apply se jen vypíše, co by se dělo.

.PARAMETER SourceSiteUrl
    Vzorový web. Jen se z něj čte.

.PARAMETER TargetSiteUrl
    Web, na který se přenáší.

.PARAMETER ClientId
    Client ID Entra ID aplikace. Když se nezadá, vezme se z PNP_CLIENT_ID.

.PARAMETER Steps
    Co přenést. Výchozí je vše: Pages, Images, HomePage, Design, Regional.

.PARAMETER ThemeName
    Název motivu pro cílový web, například "Blue" nebo název firemního motivu.
    Když se nezadá, skript zkusí motiv přečíst ze vzorového webu.

.PARAMETER Apply
    Provede změny.

.PARAMETER Pages
    Přenesou se jen tyto stránky (např. Home.aspx). Výchozí je všechny.

.EXAMPLE
    ./src/Copy-SitePages.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/P01" -TargetSiteUrl "https://contoso.sharepoint.com/sites/P42"

.EXAMPLE
    ./src/Copy-SitePages.ps1 -SourceSiteUrl "..." -TargetSiteUrl "..." -Apply

.NOTES
    Postup a řešení chyb: docs/10-stranky-a-vzhled.md

    Stránka, která v cíli existuje, se přepíše obsahem ze vzoru - u stránek
    to jinak nejde, Invoke-PnPSiteTemplate je nahrazuje celé. Ruční úpravy
    na cílové stránce se tedy ztratí. Proto ten náhled.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SourceSiteUrl,
    [Parameter(Mandatory = $true)] [string] $TargetSiteUrl,

    [string] $ClientId = $env:PNP_CLIENT_ID,
    [string] $OutputFolder = "./export",

    [ValidateSet("Pages", "Images", "HomePage", "Design", "Regional")]
    [string[]] $Steps = @("Pages", "Images", "HomePage", "Design", "Regional"),

    [string[]] $Pages = @(),

    # Název motivu, který se má nastavit na cílovém webu. Když se nezadá,
    # skript ho zkusí přečíst ze vzoru - ne každá verze PnP to ale umí.
    [string] $ThemeName = "",

    [switch] $Apply
)

$ErrorActionPreference = "Stop"

$script:Warnings = @()

function Write-Step($Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Add-PageWarning($Message) {
    $script:Warnings += $Message
    Write-Warning $Message
}

function Write-WarningSummary($OutFolder) {
    if ($script:Warnings.Count -gt 0) {
        $script:Warnings | Set-Content -Path "$OutFolder/pages-warnings.txt" -Encoding UTF8
        Write-Host ""
        Write-Host "$($script:Warnings.Count) varování - podrobnosti v $OutFolder/pages-warnings.txt" -ForegroundColor Yellow
    }
}

function Assert-Prerequisites {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "Modul PnP.PowerShell není nainstalovaný. Spusťte: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    if (-not $ClientId) {
        throw "Chybí ClientId. Zadejte -ClientId nebo nastavte PNP_CLIENT_ID."
    }
    if ($SourceSiteUrl.TrimEnd("/") -eq $TargetSiteUrl.TrimEnd("/")) {
        throw "Zdrojový a cílový web jsou tentýž."
    }
}

function Get-ServerRelativePath($SiteUrl) {
    return ([uri]$SiteUrl).AbsolutePath.TrimEnd("/")
}

# Knihovna se stránkami se jmenuje podle jazyka webu.
function Get-SitePagesList($Connection) {
    foreach ($name in @("Site Pages", "Websiteseiten", "Stránky webu")) {
        try {
            $list = Get-PnPList -Identity $name -Connection $Connection -ErrorAction Stop
            if ($list) { return $list }
        }
        catch { continue }
    }

    # Poslední pokus podle typu, ne podle názvu - ten je jazykově závislý.
    $byTemplate = Get-PnPList -Connection $Connection | Where-Object { $_.BaseTemplate -eq 119 }
    if ($byTemplate) { return @($byTemplate)[0] }

    return $null
}

function Get-PageNames($Connection) {
    $list = Get-SitePagesList $Connection
    if (-not $list) {
        Add-PageWarning "Knihovnu se stránkami se nepodařilo najít."
        return @()
    }

    $items = Get-PnPListItem -List $list.Id -Connection $Connection -PageSize 500
    return @($items | ForEach-Object { $_.FieldValues.FileLeafRef } | Where-Object { $_ -like "*.aspx" })
}

# Export-PnPPage vytvoří šablonu s jednou stránkou, Invoke-PnPSiteTemplate ji
# na cílovém webu založí. Je to stejný postup jako v CopyWebparts2 ve script.ps1.
function Copy-Page($PageName, $SourceConnection, $TargetConnection, $WorkFolder) {
    $exportPath = Join-Path $WorkFolder ("page-" + ($PageName -replace '[\\/:*?"<>|]', '_') + ".xml")

    Export-PnPPage -Force -Identity $PageName -Out $exportPath -Connection $SourceConnection -ErrorAction Stop
    Invoke-PnPSiteTemplate -Path $exportPath -WarningAction SilentlyContinue -Connection $TargetConnection -ErrorAction Stop
}

# Obrázky vložené do webpartů. Ve stránce jsou odkazem na soubor ve zdrojovém
# webu, takže se musí stáhnout a nahrát znovu, jinak v cíli chybí.
function Copy-PageImages($PageName, $SourceConnection, $TargetConnection, $SourcePath, $WorkFolder) {
    $copied = 0

    try {
        $page = Get-PnPPage -Identity $PageName -Connection $SourceConnection -ErrorAction Stop
    }
    catch {
        Add-PageWarning "Stránku '$PageName' nelze načíst pro obrázky: $($_.Exception.Message)"
        return 0
    }

    foreach ($control in $page.Controls) {
        $imageUrls = @()

        foreach ($json in @($control.ServerProcessedContent, $control.Properties)) {
            if (-not $json) { continue }
            try {
                $parsed = ConvertFrom-Json $json -ErrorAction Stop
            }
            catch { continue }

            if ($parsed.imageSources -and $parsed.imageSources.imageSource) {
                $imageUrls += @($parsed.imageSources.imageSource)
            }
        }

        foreach ($imageUrl in ($imageUrls | Where-Object { $_ } | Sort-Object -Unique)) {
            # Obrázky mimo zdrojový web (stock fotky, CDN) se nepřenášejí.
            if ($imageUrl -notlike "*$SourcePath*") { continue }

            $index = $imageUrl.IndexOf($SourcePath, [System.StringComparison]::OrdinalIgnoreCase)
            $relative = $imageUrl.Substring($index + $SourcePath.Length).TrimStart("/")
            if (-not $relative) { continue }

            $fileName = $relative.Split("/")[-1]
            $folder = $relative.Substring(0, [Math]::Max(0, $relative.Length - $fileName.Length - 1))
            if (-not $folder) { continue }

            try {
                Get-PnPFile -Url $imageUrl -Path $WorkFolder -FileName $fileName -AsFile -Force -Connection $SourceConnection -ErrorAction Stop | Out-Null
                Add-PnPFile -Path (Join-Path $WorkFolder $fileName) -Folder $folder -NewFileName $fileName -Connection $TargetConnection -ErrorAction Stop | Out-Null
                Remove-Item -Path (Join-Path $WorkFolder $fileName) -Force -ErrorAction SilentlyContinue
                $copied++
                Write-Host "      obrázek $relative" -ForegroundColor DarkGray
            }
            catch {
                Add-PageWarning "Obrázek '$relative' ze stránky '$PageName' nelze přenést: $($_.Exception.Message)"
            }
        }
    }

    return $copied
}

Assert-Prerequisites

$outputPath = $OutputFolder
New-Item -Path $outputPath -ItemType Directory -Force | Out-Null

$workFolder = Join-Path $outputPath "pages-temp"
if (Test-Path $workFolder) { Remove-Item -Path $workFolder -Recurse -Force }
New-Item -Path $workFolder -ItemType Directory -Force | Out-Null

$sourcePath = Get-ServerRelativePath $SourceSiteUrl

Write-Step "Připojuji se ke vzorovému webu"
Write-Host "  $SourceSiteUrl"
$sourceConnection = Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

Write-Step "Připojuji se k cílovému webu"
Write-Host "  $TargetSiteUrl"
$targetConnection = Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

# ============================================================
# Které stránky
# ============================================================

$pageNames = if ($Pages.Count -gt 0) { $Pages } else { Get-PageNames $sourceConnection }

Write-Step "Stránky ve vzorovém webu"
if ($pageNames.Count -eq 0) {
    Write-Host "  žádné"
}
else {
    foreach ($name in $pageNames) { Write-Host "  - $name" }
    Write-Host "  celkem: $($pageNames.Count)"
}

$existingTargetPages = @(Get-PageNames $targetConnection)
if ($existingTargetPages.Count -gt 0) {
    Write-Host "  v cíli už jsou: $($existingTargetPages -join ', ')"
}

# ============================================================
# Náhled
# ============================================================

if (-not $Apply) {
    Write-Step "NÁHLED - nic se nezmění"

    foreach ($name in $pageNames) {
        if ($Steps -notcontains "Pages") { break }
        $mark = if ($existingTargetPages -contains $name) { "přepsal bych" } else { "vytvořil bych" }
        Write-Host "  $mark  $name" -ForegroundColor Yellow
    }

    foreach ($step in @("Images", "HomePage", "Design", "Regional")) {
        if ($Steps -contains $step) { Write-Host "  provedl bych krok $step" -ForegroundColor Yellow }
    }

    if ($existingTargetPages.Count -gt 0 -and $Steps -contains "Pages") {
        Write-Host ""
        Write-Host "  Pozor: existující stránky se přepisují celé, ruční úpravy v cíli se ztratí." -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "Až bude výpis v pořádku, spusťte stejný příkaz s přepínačem -Apply." -ForegroundColor Yellow
    Write-WarningSummary $outputPath
    return
}

# ============================================================
# Stránky
# ============================================================

$pagesCopied = 0
$imagesCopied = 0

if ($Steps -contains "Pages") {
    Write-Step "Přenáším stránky"

    foreach ($name in $pageNames) {
        Write-Host "  $name"
        try {
            Copy-Page $name $sourceConnection $targetConnection $workFolder
            $pagesCopied++

            if ($Steps -contains "Images") {
                $imagesCopied += Copy-PageImages $name $sourceConnection $targetConnection $sourcePath $workFolder
            }
        }
        catch {
            Add-PageWarning "Stránku '$name' nelze přenést: $($_.Exception.Message)"
        }
    }

    Write-Host "  přeneseno stránek: $pagesCopied, obrázků: $imagesCopied"
}

# ============================================================
# Domovská stránka
# ============================================================

if ($Steps -contains "HomePage") {
    Write-Step "Nastavuji domovskou stránku"
    try {
        $homePage = Get-PnPHomePage -Connection $sourceConnection -ErrorAction Stop
        Set-PnPHomePage -RootFolderRelativeUrl $homePage -Connection $targetConnection -ErrorAction Stop
        Write-Host "  $homePage" -ForegroundColor Green
    }
    catch {
        Add-PageWarning "Domovskou stránku nelze nastavit: $($_.Exception.Message)"
    }
}

# ============================================================
# Vzhled webu
# ============================================================

if ($Steps -contains "Design") {
    Write-Step "Přenáším vzhled webu"

    # WebSettings nese logo, popis a pár vlastností webu, ale NENESE barevné
    # téma moderního webu - to se nastavuje zvlášť přes Set-PnPWebTheme.
    $designPath = Join-Path $workFolder "websettings.xml"
    try {
        Get-PnPSiteTemplate -Out $designPath -Handlers WebSettings -Force -Connection $sourceConnection -ErrorAction Stop
        Invoke-PnPSiteTemplate -Path $designPath -WarningAction SilentlyContinue -Connection $targetConnection -ErrorAction Stop
        Write-Host "  WebSettings přeneseny" -ForegroundColor Green
    }
    catch {
        Add-PageWarning "WebSettings nelze přenést: $($_.Exception.Message)"
    }

    # Barevné téma. Rozhoduje o barvách webu a ve WebSettings NENÍ.
    # Cmdlet na přečtení tématu ze zdroje se mezi verzemi PnP liší a v některých
    # chybí úplně - proto se dá název tématu předat parametrem -ThemeName.
    $themeName = $ThemeName

    if (-not $themeName) {
        $reader = @("Get-PnPWebTheme", "Get-PnPTheme") |
            Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } |
            Select-Object -First 1

        if ($reader) {
            try {
                $sourceTheme = & $reader -Connection $sourceConnection -ErrorAction Stop

                # Cmdlet vrací objekt; když v něm název není, nemá smysl brát
                # jeho ToString() - to je jen jméno typu.
                foreach ($property in @("Name", "Theme", "ThemeName")) {
                    if ($sourceTheme.$property -and "$($sourceTheme.$property)" -notlike "PnP.*") {
                        $themeName = "$($sourceTheme.$property)"
                        break
                    }
                }

                if (-not $themeName) {
                    Add-PageWarning "Motiv vzoru se načetl, ale nemá pojmenovaný motiv (je to vlastní nastavení barev). Zjistěte název v cílovém webu a předejte ho jako -ThemeName, nebo klíčem themeName v konfiguraci."
                }
            }
            catch {
                Add-PageWarning "Téma vzoru nelze přečíst ($reader): $($_.Exception.Message)"
            }
        }
        else {
            Add-PageWarning "Nainstalovaná verze PnP.PowerShell neumí přečíst téma webu. Zjistěte název tématu v cílovém webu (Nastavení -> Změnit vzhled -> Motiv) a předejte ho jako -ThemeName, nebo v konfiguraci klíčem themeName."
        }
    }

    if ($themeName) {
        if (Get-Command Set-PnPWebTheme -ErrorAction SilentlyContinue) {
            try {
                Set-PnPWebTheme -Theme $themeName -Connection $targetConnection -ErrorAction Stop
                Write-Host "  téma: $themeName" -ForegroundColor Green
            }
            catch {
                $detail = $_.Exception.Message
                $hint = if ($detail -like "*unauthorized*") {
                    "Na nastavení motivu nemáte oprávnění - potřebujete být vlastníkem webu."
                } else {
                    "Vlastní motiv musí být registrovaný v tenantu (Add-PnPTenantTheme)."
                }
                Add-PageWarning "Motiv '$themeName' nelze nastavit: $detail $hint"
            }
        }
        else {
            Add-PageWarning "Set-PnPWebTheme v nainstalované verzi PnP.PowerShell není, téma nastavte ručně."
        }
    }

    # Hlavička, megamenu a levá navigace - další věci, které jsou vidět na
    # první pohled a WebSettings je nenese. Každá zvlášť, aby jedna
    # nepodporovaná vlastnost neshodila ostatní.
    $sourceWeb = $null
    try {
        $sourceWeb = Get-PnPWeb -Connection $sourceConnection `
            -Includes HeaderLayout, HeaderEmphasis, MegaMenuEnabled, QuickLaunchEnabled
    }
    catch {
        Add-PageWarning "Nastavení hlavičky ze vzoru nelze přečíst: $($_.Exception.Message)"
    }

    if ($sourceWeb) {
        if ($sourceWeb.HeaderLayout) {
            try {
                Set-PnPWeb -HeaderLayout $sourceWeb.HeaderLayout -Connection $targetConnection -ErrorAction Stop
                Write-Host "  HeaderLayout: $($sourceWeb.HeaderLayout)" -ForegroundColor Green
            }
            catch { Add-PageWarning "HeaderLayout nelze nastavit: $($_.Exception.Message)" }
        }

        if ($sourceWeb.HeaderEmphasis) {
            try {
                Set-PnPWeb -HeaderEmphasis $sourceWeb.HeaderEmphasis -Connection $targetConnection -ErrorAction Stop
                Write-Host "  HeaderEmphasis: $($sourceWeb.HeaderEmphasis)" -ForegroundColor Green
            }
            catch { Add-PageWarning "HeaderEmphasis nelze nastavit: $($_.Exception.Message)" }
        }

        try {
            Set-PnPWeb -MegaMenuEnabled:([bool]$sourceWeb.MegaMenuEnabled) -Connection $targetConnection -ErrorAction Stop
            Write-Host "  MegaMenu: $([bool]$sourceWeb.MegaMenuEnabled)" -ForegroundColor Green
        }
        catch { Add-PageWarning "MegaMenu nelze nastavit: $($_.Exception.Message)" }

        try {
            Set-PnPWeb -QuickLaunchEnabled:([bool]$sourceWeb.QuickLaunchEnabled) -Connection $targetConnection -ErrorAction Stop
            Write-Host "  QuickLaunch: $([bool]$sourceWeb.QuickLaunchEnabled)" -ForegroundColor Green
        }
        catch { Add-PageWarning "QuickLaunch nelze nastavit: $($_.Exception.Message)" }
    }
}

# ============================================================
# Regionální nastavení
# ============================================================

if ($Steps -contains "Regional") {
    Write-Step "Přenáším regionální nastavení"

    $regionalPath = Join-Path $workFolder "regional.xml"
    try {
        Get-PnPSiteTemplate -Out $regionalPath -Handlers RegionalSettings -Force -Connection $sourceConnection -ErrorAction Stop
        Invoke-PnPSiteTemplate -Path $regionalPath -WarningAction SilentlyContinue -Connection $targetConnection -ErrorAction Stop
        Write-Host "  přeneseno" -ForegroundColor Green
    }
    catch {
        Add-PageWarning "Regionální nastavení nelze přenést: $($_.Exception.Message). Původní script.ps1 slučuje jen vybrané atributy - viz docs/10-stranky-a-vzhled.md"
    }
}

Write-WarningSummary $outputPath
Write-Step "Hotovo"
