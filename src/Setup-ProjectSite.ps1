<#
.SYNOPSIS
    Připraví nový projektový web: strukturu složek z Excelu, seznamy a kalendáře
    ze vzorového webu. Jeden vstupní bod pro všechno ostatní.

.DESCRIPTION
    Obalává tři skripty, které dělají samotnou práci:

      New-FolderStructure.ps1     složky z Excelu + metadata
      Copy-SharePointLists.ps1    seznamy ze vzorového webu
      Copy-SharePointEvents.ps1   kalendáře (Events) ze vzorového webu
      Copy-SitePages.ps1          stránky, obrázky, vzhled, regionální nastavení
      Set-CsdClass.ps1            CSD Class na knihovně a jejích souborech
      Copy-SiteNavigation.ps1     navigace (volitelně)
      script.ps1                  původní klonovací skript (volitelně)

    Všechny leží v src/ vedle tohoto skriptu.

    Nastavení, které se skoro nemění - ClientId, vzorový web, knihovna, hodnota
    CSD Class - se drží v config/settings.json, takže se nepíše do příkazu.

    BEZ PŘEPÍNAČE -Apply SE NIC NEZMĚNÍ. Vypíše se jen, co by se dělalo.

.PARAMETER TargetSiteUrl
    Web, který se má připravit. Jediná věc, která se mění při každém spuštění.

.PARAMETER Apply
    Provede změny. Bez něj jen výpis.

.PARAMETER WithData
    Seznamy a kalendáře se přenesou i s položkami. Bez tohoto přepínače se
    přenese jen jejich struktura - to je běžnější případ.

.PARAMETER Steps
    Co se má dělat. Výchozí: Folders, Lists, Events, Pages, Navigation, CsdClass.
    TemplateClone je potřeba vyžádat výslovně.

    TemplateClone spustí původní script.ps1, který naklonuje vzorový web jako
    celek. MAŽE v cíli seznamy, než je vytvoří znovu - použitelné jen na
    čerstvě založený web. Podrobněji v docs/11-klonovani-vzoru.md.

.PARAMETER SourceSiteUrl
    Vzorový web. Normálně se bere z konfigurace, tímto se dá jednorázově změnit.

.PARAMETER ConfigPath
    Cesta ke konfiguraci. Výchozí config/settings.json.

.EXAMPLE
    # 1) Podívat se, co by se stalo
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42"

.EXAMPLE
    # 2) Připravit web (složky, seznamy a kalendáře bez dat)
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42" -Apply

.EXAMPLE
    # Včetně obsahu seznamů a kalendářů
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42" -WithData -Apply

.EXAMPLE
    # Naklonovat vzorový web jako celek a pak doplnit složky z Excelu
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42" -Steps TemplateClone,Folders,CsdClass -Apply

.EXAMPLE
    # Jen stránky a vzhled, nic jiného
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42" -Steps Pages -Apply

.EXAMPLE
    # Jen složky
    ./src/Setup-ProjectSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/Proj42" -Steps Folders -Apply

.NOTES
    Postup: docs/09-jeden-skript.md
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $TargetSiteUrl,

    [switch] $Apply,
    [switch] $WithData,

    [ValidateSet("TemplateClone", "Folders", "Lists", "Events", "Pages", "Navigation", "CsdClass")]
    [string[]] $Steps = @("Folders", "Lists", "Events", "Pages", "Navigation", "CsdClass"),

    [string] $SourceSiteUrl = "",
    [string] $ConfigPath = "",
    [string] $OutputFolder = "./export"
)

$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptRoot

function Write-Header($Message) {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host " $Message" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
}

function Write-Step($Message) {
    Write-Host ""
    Write-Host "--- $Message" -ForegroundColor Cyan
}

function Resolve-RepoPath($Path) {
    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return Join-Path $repoRoot $Path
}

function Import-Settings($Path) {
    if (-not $Path) { $Path = Join-Path $repoRoot "config/settings.json" }

    if (-not (Test-Path $Path)) {
        $example = Join-Path $repoRoot "config/settings.example.json"
        throw @"
Chybí konfigurace: $Path

Vytvořte ji zkopírováním šablony a doplněním hodnot:

    Copy-Item "$example" "$Path"

Pak v ní vyplňte alespoň clientId a sourceSiteUrl. Do repozitáře se nedostane,
je v .gitignore. Podrobněji v docs/09-jeden-skript.md
"@
    }

    try {
        return Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Konfiguraci '$Path' nelze přečíst - není to platný JSON? $($_.Exception.Message)"
    }
}

# ConvertFrom-Json vrací PSCustomObject, ale -FixedMetadata chce hashtable.
function ConvertTo-Hashtable($Object) {
    $table = @{}
    if (-not $Object) { return $table }

    foreach ($property in $Object.PSObject.Properties) {
        if ($property.Name.StartsWith("_")) { continue }
        $table[$property.Name] = $property.Value
    }
    return $table
}

function Assert-ScriptsPresent($Names) {
    foreach ($name in $Names) {
        $path = Join-Path $scriptRoot $name
        if (-not (Test-Path $path)) {
            throw "Chybí skript $name ve složce src/. Máte kompletní repozitář?"
        }
    }
}

$results = [ordered]@{}

function Invoke-Step($Name, $Action) {
    if ($Steps -notcontains $Name) { return }

    Write-Step $Name
    try {
        & $Action
        if (-not $results.Contains($Name)) { $results[$Name] = "OK" }
    }
    catch {
        $results[$Name] = "CHYBA: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "Krok '$Name' selhal: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Pokračuji dalšími kroky." -ForegroundColor Yellow
    }
}

# ============================================================
# Načtení nastavení
# ============================================================

$settings = Import-Settings $ConfigPath

$clientId = $settings.clientId
if (-not $clientId -or $clientId -like "0000*") {
    throw "V konfiguraci není vyplněné clientId. Doplňte ho do config/settings.json."
}

if (-not $SourceSiteUrl) { $SourceSiteUrl = $settings.sourceSiteUrl }

$library = if ($settings.library) { $settings.library } else { "Dokumenty" }
$structureFile = Resolve-RepoPath $(if ($settings.folderStructureFile) { $settings.folderStructureFile } else { "Folder_Structure.xlsx" })
$fixedMetadata = ConvertTo-Hashtable $settings.fixedMetadata

$needsSource = @("TemplateClone", "Lists", "Events", "Pages", "Navigation") | Where-Object { $Steps -contains $_ }
if ($needsSource -and -not $SourceSiteUrl) {
    throw "Kroky $($needsSource -join ', ') potřebují vzorový web. Doplňte sourceSiteUrl do konfigurace, nebo předejte -SourceSiteUrl."
}

New-Item -Path (Resolve-RepoPath $OutputFolder) -ItemType Directory -Force | Out-Null

# ============================================================
# Přehled před spuštěním
# ============================================================

Write-Header $(if ($Apply) { "PŘÍPRAVA WEBU" } else { "NÁHLED - nic se nezmění" })

Write-Host " cílový web:     $TargetSiteUrl"
if ($SourceSiteUrl) { Write-Host " vzorový web:    $SourceSiteUrl" }
Write-Host " knihovna:       $library"
Write-Host " struktura:      $structureFile"
if ($fixedMetadata.Count -gt 0) {
    $metaText = (($fixedMetadata.GetEnumerator() | Sort-Object Name |
                  ForEach-Object { "$($_.Key) = $($_.Value)" }) -join "; ")
    Write-Host " metadata:       $metaText"
}
Write-Host " kroky:          $($Steps -join ', ')"
Write-Host " data v listech: $(if ($WithData) { 'ANO - přenesou se i položky' } else { 'ne - jen struktura' })"

if (-not $Apply) {
    Write-Host ""
    Write-Host " Režim náhledu. Pro provedení přidejte -Apply" -ForegroundColor Yellow
}

Assert-ScriptsPresent @(
    "New-FolderStructure.ps1",
    "Set-CsdClass.ps1",
    "Copy-SharePointLists.ps1",
    "Copy-SharePointEvents.ps1",
    "Copy-SitePages.ps1",
    "Copy-SiteNavigation.ps1"
)

# ============================================================
# Klonování vzorového webu původním script.ps1
#
# script.ps1 nemá parametry - konfigurace je napsaná v jeho hlavičce. Aby se
# nemusel upravovat (je společný a používá ho i byznys ručně), vygeneruje se
# jeho kopie s doplněnými hodnotami a spustí se ta. Originál zůstane nedotčený.
# ============================================================

Invoke-Step "TemplateClone" {
    $legacyPath = Join-Path $scriptRoot "script.ps1"
    if (-not (Test-Path $legacyPath)) {
        throw "src/script.ps1 chybí."
    }

    $options = $settings.legacyScript
    $domain = ([uri]$TargetSiteUrl).GetLeftPart([System.UriPartial]::Authority)
    $sourceDomain = ([uri]$SourceSiteUrl).GetLeftPart([System.UriPartial]::Authority)

    if ($sourceDomain -ne $domain) {
        throw "script.ps1 umí kopírovat jen v rámci jednoho tenantu, ale vzor je na $sourceDomain a cíl na $domain."
    }

    # CopyCount = 1 znamená, že se nezpracuje žádný seznam - podmínka ve
    # script.ps1 je "$counter -lt $CopyCount" a counter začíná na 1. Stránky,
    # navigace a vzhled se přenesou i tak, volají se až za tou smyčkou.
    $arguments = @{
        SiteDomain             = $domain
        SourcePath             = ([uri]$SourceSiteUrl).AbsolutePath.TrimEnd("/")
        TargetPath             = ([uri]$TargetSiteUrl).AbsolutePath.TrimEnd("/")
        ClientId               = $clientId
        CopyCount              = $(if ($options.copyLists -eq $false) { 1 } else { 1000 })
        SetOfflineAvailable    = $(if ($options.setOfflineAvailable -eq $true) { "ja" } else { "nein" })
        IsCopyPages            = ($options.copyPages -ne $false)
        IsCopyTemplateDesign   = ($options.copyDesign -ne $false)
        IsCopyRegionalSettings = ($options.copyRegional -ne $false)
        IsCopyNavigation       = ($options.copyNavigation -ne $false)
    }

    Write-Host "  vzor:     $($arguments.SourcePath)"
    Write-Host "  cíl:      $($arguments.TargetPath)"
    Write-Host "  seznamy:  $(if ($arguments.CopyCount -eq 1) { 'ne' } else { 'ano' })"
    Write-Host "  stránky:  $($arguments.IsCopyPages)   vzhled: $($arguments.IsCopyTemplateDesign)   navigace: $($arguments.IsCopyNavigation)"

    if (-not $Apply) {
        Write-Host ""
        Write-Host "  [náhled] spustil bych script.ps1 s těmito parametry. Režim náhledu nemá." -ForegroundColor Yellow
        Write-Host "  POZOR: script.ps1 v cíli MAŽE seznamy, než je vytvoří znovu." -ForegroundColor Red
        Write-Host "  Používejte ho jen na čerstvě založený web." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "  Spouštím klonování. Maže a znovu vytváří seznamy v cíli." -ForegroundColor Yellow
    & $legacyPath @arguments
}

# ============================================================
# Složky z Excelu
# ============================================================

Invoke-Step "Folders" {
    $arguments = @{
        SiteUrl       = $TargetSiteUrl
        Library       = $library
        Path          = $structureFile
        ClientId      = $clientId
        OutputFolder  = (Resolve-RepoPath $OutputFolder)
        FixedMetadata = $fixedMetadata
    }
    if ($Apply) { $arguments["Apply"] = $true }

    & (Join-Path $scriptRoot "New-FolderStructure.ps1") @arguments
}

# ============================================================
# CSD Class na knihovně
#
# Metadata na složce platí pro složku, ne pro soubory v ní. Aby hodnotu dostal
# každý nově nahraný soubor, musí být nastavená jako výchozí hodnota sloupce
# v knihovně - to je jiný mechanismus než zápis na položku.
# ============================================================

Invoke-Step "CsdClass" {
    if ($fixedMetadata.Count -eq 0) {
        Write-Host "  žádná konstantní metadata v konfiguraci, nic k nastavení"
        $results["CsdClass"] = "nic k nastavení"
        return
    }

    $scope = if ($settings.csdClassScope) { $settings.csdClassScope } else { "DefaultValue" }

    foreach ($field in ($fixedMetadata.Keys | Sort-Object)) {
        $arguments = @{
            SiteUrl  = $TargetSiteUrl
            Library  = $library
            Field    = $field
            Value    = $fixedMetadata[$field]
            Scope    = $scope
            ClientId = $clientId
        }

        # Popisek sloupce, pod kterým ho uvidí uživatel v knihovně.
        if ($settings.fieldTitles -and $settings.fieldTitles.$field) {
            $arguments["Title"] = $settings.fieldTitles.$field
        }
        if ($Apply) { $arguments["Apply"] = $true }

        & (Join-Path $scriptRoot "Set-CsdClass.ps1") @arguments
    }
}

# ============================================================
# Seznamy a kalendáře
#
# Tyto skripty nemají režim náhledu - buď se spustí, nebo ne. V náhledu se
# proto jen vypíše, co by se spustilo.
# ============================================================

$copyValues = if ($WithData) { "yes" } else { "no" }

# Weby, jejichž navigace nebo dlaždice se plní ze seznamu, budou bez -WithData
# vypadat prázdné - struktura seznamu vznikne, ale položky v ní ne.
if (-not $WithData -and ($Steps -contains "Lists")) {
    Write-Host ""
    Write-Host " Poznámka: seznamy se přenesou bez položek. Pokud vzorový web plní" -ForegroundColor DarkGray
    Write-Host " navigaci nebo dlaždice ze seznamu (Navigation, Hyperlinks), budou" -ForegroundColor DarkGray
    Write-Host " v cíli prázdné. V takovém případě použijte -WithData." -ForegroundColor DarkGray
}

Invoke-Step "Lists" {
    if (-not $Apply) {
        Write-Host "  [náhled] spustil bych Copy-SharePointLists.ps1, data: $copyValues" -ForegroundColor Yellow
        Write-Host "  Tento skript nemá režim náhledu - v cíli zakládá chybějící seznamy." -ForegroundColor DarkGray
        return
    }

    & (Join-Path $scriptRoot "Copy-SharePointLists.ps1") `
        -SourceSiteUrl $SourceSiteUrl `
        -TargetSiteUrl $TargetSiteUrl `
        -CopyValues $copyValues `
        -ClientId $clientId
}

Invoke-Step "Events" {
    if (-not $Apply) {
        Write-Host "  [náhled] spustil bych Copy-SharePointEvents.ps1, data: $copyValues" -ForegroundColor Yellow
        Write-Host "  Tento skript nemá režim náhledu - v cíli zakládá chybějící kalendáře." -ForegroundColor DarkGray
        return
    }

    & (Join-Path $scriptRoot "Copy-SharePointEvents.ps1") `
        -SourceSiteUrl $SourceSiteUrl `
        -TargetSiteUrl $TargetSiteUrl `
        -CopyValues $copyValues `
        -ClientId $clientId
}

# ============================================================
# Stránky, obrázky, vzhled a regionální nastavení
# ============================================================

Invoke-Step "Pages" {
    $arguments = @{
        SourceSiteUrl = $SourceSiteUrl
        TargetSiteUrl = $TargetSiteUrl
        ClientId      = $clientId
        OutputFolder  = (Resolve-RepoPath $OutputFolder)
    }
    if ($Apply) { $arguments["Apply"] = $true }

    & (Join-Path $scriptRoot "Copy-SitePages.ps1") @arguments
}

# ============================================================
# Navigace (jen na vyžádání)
# ============================================================

Invoke-Step "Navigation" {
    $arguments = @{
        SourceSiteUrl = $SourceSiteUrl
        TargetSiteUrl = $TargetSiteUrl
        ClientId      = $clientId
        OutputFolder  = (Resolve-RepoPath $OutputFolder)
    }
    if ($Apply) { $arguments["Apply"] = $true }

    & (Join-Path $scriptRoot "Copy-SiteNavigation.ps1") @arguments
}

# ============================================================
# Souhrn
# ============================================================

Write-Header "SOUHRN"

foreach ($key in $results.Keys) {
    $value = $results[$key]
    $color = if ($value -like "CHYBA*") { "Red" } elseif ($value -eq "OK") { "Green" } else { "Yellow" }
    Write-Host (" {0,-15} {1}" -f $key, $value) -ForegroundColor $color
}

$skipped = @($Steps | Where-Object { -not $results.Contains($_) })
if ($skipped.Count -gt 0) {
    Write-Host (" {0,-15} {1}" -f "neproběhlo", ($skipped -join ", ")) -ForegroundColor DarkGray
}

Write-Host ""
if (-not $Apply) {
    Write-Host " Nic se nezměnilo. Pro provedení spusťte stejný příkaz s -Apply" -ForegroundColor Yellow
}
else {
    Write-Host " Výstupy a plány najdete v $(Resolve-RepoPath $OutputFolder)"
}

$failed = @($results.Values | Where-Object { $_ -like "CHYBA*" })
if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host " $($failed.Count) krok(ů) selhalo - viz červené řádky výše." -ForegroundColor Red
    exit 1
}
