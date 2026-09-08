<#
.SYNOPSIS
    Vytvoří v knihovně dokumentů na SharePointu strukturu složek podle Excelu
    nebo CSV, včetně metadat u jednotlivých složek.

.DESCRIPTION
    Vstupní tabulka má jeden řádek na složku. Cesta se skládá ze sloupců
    Level1..LevelN - použijí se všechny neprázdné zleva:

        Level1            Level2      Level3     -> vytvořená cesta
        01_Organisation                          -> 01_Organisation
        01_Organisation   04_RASI                -> 01_Organisation/04_RASI
        04_Meetings       07_Groups   05_Zones   -> 04_Meetings/07_Groups/05_Zones

    Ostatní sloupce se zapíší jako metadata té složky, o které řádek je.
    Které kam, určuje -MetadataMap; výchozí hodnota odpovídá tabulce
    Folder_Structure.xlsx.

    VÝCHOZÍ REŽIM JE DRY-RUN. Bez přepínače -Apply se nic nezapíše, jen se
    vypíše a uloží plán. Skript nikdy nic nemaže - existující složky přeskočí,
    chybějící doplní, takže se dá pouštět opakovaně.

.PARAMETER SiteUrl
    Plná URL webu, kde se má struktura vytvořit.

.PARAMETER Library
    Cílová knihovna dokumentů. Přijímá GUID, název, nebo cestu relativní k webu
    ("Shared Documents"). GUID je nejspolehlivější - názvy výchozích knihoven
    jsou na serveru lokalizované.

.PARAMETER Path
    Cesta k .xlsx nebo .csv se strukturou.

.PARAMETER ClientId
    Client ID Entra ID aplikace. Když se nezadá, vezme se z PNP_CLIENT_ID.

.PARAMETER Apply
    Provede změny. Bez něj skript jen vypíše, co by udělal.

.PARAMETER SkipMetadata
    Vytvoří jen složky - metadata neřeší a sloupce nezakládá.

.PARAMETER MetadataMap
    Mapování sloupců tabulky na sloupce v SharePointu. Klíč je hlavička
    v tabulce, hodnota popisuje cílový sloupec.

.PARAMETER FixedMetadata
    Konstantní metadata vyražená na každou vytvořenou složku - klíč je interní
    název sloupce, hodnota je text. Sloupec, který v knihovně chybí, se vytvoří
    jako Text. Když stejný sloupec plní i tabulka, vyhrává hodnota z tabulky.

.EXAMPLE
    # 1) Nejdřív se podívat, co by se stalo
    ./src/New-FolderStructure.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Proj01" -Library "Shared Documents" -Path ./Folder_Structure.xlsx

.EXAMPLE
    # Každé složce navíc nastavit CSD na pevnou hodnotu
    ./src/New-FolderStructure.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Proj01" -Library "Shared Documents" -Path ./Folder_Structure.xlsx -FixedMetadata @{ CSD = "5.3 Car Series and Concept Docs" } -Apply

.EXAMPLE
    # 2) Teprve pak vytvořit
    ./src/New-FolderStructure.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Proj01" -Library "Shared Documents" -Path ./Folder_Structure.xlsx -Apply

.NOTES
    Postup a řešení chyb: docs/07-struktura-slozek-z-excelu.md

    Čtení .xlsx vyžaduje modul ImportExcel:
        Install-Module ImportExcel -Scope CurrentUser
    Bez něj tabulku v Excelu uložte jako CSV a předejte .csv - to funguje bez
    instalace čehokoli dalšího.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $Library,
    [Parameter(Mandatory = $true)] [string] $Path,

    [string] $ClientId = $env:PNP_CLIENT_ID,
    [string] $OutputFolder = "./export",

    [switch] $Apply,
    [switch] $SkipMetadata,

    # Prázdné = autodetekce podle hlavičky. České Excely ukládají CSV se ";".
    [string] $Delimiter = "",

    [hashtable] $MetadataMap = @{
        "Made in/responsible" = @{ InternalName = "Responsible"; DisplayName = "Responsible"     }
        "English translation" = @{ InternalName = "NameEnglish"; DisplayName = "Name (English)"  }
        "German translation"  = @{ InternalName = "NameGerman";  DisplayName = "Name (Deutsch)"  }
    },

    # Konstantní metadata vyražená na každou vytvořenou složku. Klíč je interní
    # název sloupce v SharePointu, hodnota je text, který se do něj zapíše.
    # Příklad: -FixedMetadata @{ CSD = "5.3 Car Series and Concept Docs" }
    # Když stejný sloupec plní i tabulka, hodnota z tabulky má přednost.
    [hashtable] $FixedMetadata = @{}
)

$ErrorActionPreference = "Stop"

$script:Warnings = @()

function Write-Step($Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Add-StructureWarning($Message) {
    $script:Warnings += $Message
    Write-Warning $Message
}

function Assert-Prerequisites {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "Modul PnP.PowerShell není nainstalovaný. Spusťte: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    if (-not $ClientId) {
        throw "Chybí ClientId. Zadejte -ClientId, nebo nastavte proměnnou PNP_CLIENT_ID. Podrobněji v docs/06-jak-spustit-export.md"
    }
    if (-not (Test-Path $Path)) {
        throw "Vstupní soubor '$Path' neexistuje."
    }
}

function Import-StructureTable($TablePath, $CsvDelimiter) {
    $extension = [System.IO.Path]::GetExtension($TablePath).ToLowerInvariant()

    if ($extension -eq ".csv") {
        if (-not $CsvDelimiter) {
            $header = Get-Content -Path $TablePath -TotalCount 1
            $semicolons = @($header.ToCharArray() | Where-Object { $_ -eq ";" }).Count
            $commas = @($header.ToCharArray() | Where-Object { $_ -eq "," }).Count
            $CsvDelimiter = if ($semicolons -ge $commas) { ";" } else { "," }
            Write-Host "  oddělovač CSV: '$CsvDelimiter' (autodetekce)"
        }
        return Import-Csv -Path $TablePath -Delimiter $CsvDelimiter
    }

    if ($extension -eq ".xlsx" -or $extension -eq ".xlsm" -or $extension -eq ".xls") {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw @"
Pro čtení souboru $extension je potřeba modul ImportExcel:

    Install-Module ImportExcel -Scope CurrentUser

Nebo tabulku v Excelu uložte jako CSV (Soubor > Uložit jako > CSV UTF-8) a
předejte skriptu ten .csv - to funguje bez instalace čehokoli.
"@
        }
        Import-Module ImportExcel
        return Import-Excel -Path $TablePath
    }

    throw "Nepodporovaný formát '$extension'. Použijte .xlsx nebo .csv."
}

# Sloupce s úrovněmi zanoření. Bere všechny se jménem Level<číslo>, takže
# tabulka může mít libovolnou hloubku, ne jen čtyři úrovně.
function Get-LevelColumnNames($Rows) {
    $names = $Rows[0].PSObject.Properties.Name |
        Where-Object { $_ -match '^\s*Level\s*\d+\s*$' } |
        Sort-Object { [int]([regex]::Match($_, '\d+').Value) }

    if (-not $names) {
        $found = ($Rows[0].PSObject.Properties.Name -join ", ")
        throw "V tabulce nejsou sloupce Level1, Level2, ... Nalezené sloupce: $found"
    }
    return @($names)
}

function Test-FolderNameValid($Name) {
    if ($Name -match '["*:<>?/\\|]') { return $false }
    if ($Name -ne $Name.Trim()) { return $false }
    if ($Name.EndsWith(".")) { return $false }
    if ($Name -eq "forms") { return $false }
    return $true
}

# Z tabulky udělá seznam složek s cestou, hloubkou a metadaty. Chybějící
# mezilehlé složky doplní, i když pro ně v tabulce vlastní řádek není.
function Get-DesiredFolders($Rows, $LevelColumns, $MetaMap, $Fixed) {
    $folders = [ordered]@{}

    foreach ($row in $Rows) {
        $segments = @()
        $gapFound = $false
        $malformed = $false

        foreach ($column in $LevelColumns) {
            $value = "$($row.$column)".Trim()

            if (-not $value) { $gapFound = $true; continue }
            # Zaplněná úroveň za prázdnou znamená rozbitý řádek.
            if ($gapFound) { $malformed = $true; break }
            $segments += $value
        }

        if ($malformed) {
            $dump = (@($LevelColumns | ForEach-Object { $row.$_ }) -join " | ")
            Add-StructureWarning "Přeskakuji řádek s dírou v úrovních: $dump"
            continue
        }
        if ($segments.Count -eq 0) { continue }

        $invalid = @($segments | Where-Object { -not (Test-FolderNameValid $_) })
        if ($invalid.Count -gt 0) {
            Add-StructureWarning "Přeskakuji '$($segments -join '/')' - nepovolený název složky: $($invalid -join ', ')"
            continue
        }

        for ($i = 1; $i -le $segments.Count; $i++) {
            $path = ($segments[0..($i - 1)]) -join "/"
            if (-not $folders.Contains($path)) {
                # Vlastní kopie pro každou složku, ne odkaz na společnou tabulku.
                $metadata = @{}
                foreach ($fixedKey in $Fixed.Keys) {
                    $metadata[$fixedKey] = "$($Fixed[$fixedKey])"
                }

                $folders[$path] = [pscustomobject]@{
                    Path     = $path
                    Name     = $segments[$i - 1]
                    Depth    = $i
                    Metadata = $metadata
                    Implicit = $true
                }
            }
        }

        # Metadata patří jen k té složce, o které řádek je.
        $leaf = $folders[($segments -join "/")]
        $leaf.Implicit = $false
        foreach ($sourceColumn in $MetaMap.Keys) {
            $value = "$($row.$sourceColumn)".Trim()
            if ($value) {
                $leaf.Metadata[$MetaMap[$sourceColumn].InternalName] = $value
            }
        }
    }

    return @($folders.Values | Sort-Object Depth, Path)
}

# Knihovnu hledá podle GUIDu, názvu i cesty. Systémové knihovny vynechává.
function Resolve-Library($Identity) {
    $allLists = Get-PnPList -Includes RootFolder

    $match = $allLists | Where-Object { $_.Id.ToString() -eq $Identity }
    if (-not $match) { $match = $allLists | Where-Object { $_.Title -eq $Identity } }
    if (-not $match) {
        $match = $allLists | Where-Object {
            $_.RootFolder.ServerRelativeUrl -eq $Identity -or
            $_.RootFolder.ServerRelativeUrl.EndsWith("/$Identity")
        }
    }

    if (-not $match) {
        $systemLibs = @("Site Assets", "Websiteobjekte", "Style Library", "Form Templates", "Site Pages", "Websiteseiten")
        $available = ($allLists |
            Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden -and $_.Title -notin $systemLibs } |
            ForEach-Object { "  - $($_.Title)   [$($_.Id)]" }) -join "`n"
        throw "Knihovnu '$Identity' jsem nenašel. Dostupné knihovny dokumentů:`n$available"
    }

    return @($match)[0]
}

function Get-SiteRelativeUrl($ServerRelativeUrl, $SiteRoot) {
    $relative = $ServerRelativeUrl
    if ($SiteRoot -and $SiteRoot -ne "/" -and $relative.StartsWith($SiteRoot)) {
        $relative = $relative.Substring($SiteRoot.Length)
    }
    return $relative.TrimStart("/")
}

# Jedním CAML dotazem získá všechny existující složky v knihovně i s jejich
# item ID. Slouží zároveň ke zjištění, co existuje, i k zápisu metadat.
function Get-ExistingFolderMap($List, $LibraryRoot) {
    $caml = "<View Scope='RecursiveAll'><Query><Where><Eq>" +
            "<FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value>" +
            "</Eq></Where></Query></View>"

    $map = @{}
    foreach ($item in (Get-PnPListItem -List $List.Id -Query $caml -PageSize 2000)) {
        $fileRef = $item.FieldValues.FileRef
        if (-not $fileRef) { continue }

        $relative = $fileRef
        if ($relative.StartsWith($LibraryRoot)) {
            $relative = $relative.Substring($LibraryRoot.Length)
        }
        $relative = $relative.TrimStart("/")
        if ($relative) { $map[$relative] = $item.Id }
    }
    return $map
}

# Sloupce, které musí v knihovně existovat: z -MetadataMap i z -FixedMetadata.
function Get-TargetColumns($MetaMap, $Fixed) {
    $targets = [ordered]@{}

    foreach ($key in $MetaMap.Keys) {
        $target = $MetaMap[$key]
        $targets[$target.InternalName] = [pscustomobject]@{
            InternalName = $target.InternalName
            DisplayName  = $target.DisplayName
        }
    }

    foreach ($key in $Fixed.Keys) {
        if (-not $targets.Contains($key)) {
            $targets[$key] = [pscustomobject]@{
                InternalName = $key
                DisplayName  = $key
            }
        }
    }

    return @($targets.Values)
}

# Mapa pro překlad na interní názvy sloupců. Set-PnPListItem přijímá interní
# název, ale uživatel zadává ten, který vidí v SharePointu - a ten se u sloupce
# s mezerou liší ("CSD Class" -> "CSD_x0020_Class"). PowerShellové hashtable
# porovnávají klíče bez ohledu na velikost písmen, takže stačí jedna mapa.
function Get-FieldNameMap($List) {
    $map = @{}
    foreach ($field in (Get-PnPField -List $List.Id)) {
        foreach ($alias in @($field.InternalName, $field.StaticName, $field.Title)) {
            if ($alias -and -not $map.ContainsKey($alias)) {
                $map[$alias] = $field.InternalName
            }
        }
    }
    return $map
}

# Interní název pro nově zakládaný sloupec. Mezery a diakritika by se zakódovaly
# do nečitelného _x0020_, takže je rovnou vynecháme a hezký název dáme do titulku.
function Get-SafeInternalName($Name) {
    $clean = ($Name -replace '[^A-Za-z0-9_]', '')
    if (-not $clean) { throw "Z názvu sloupce '$Name' nelze odvodit interní název." }
    if ($clean -match '^\d') { $clean = "f$clean" }
    return $clean
}

function Convert-MetadataKeys($Metadata, $FieldMap, $FolderPath) {
    $converted = @{}
    foreach ($key in $Metadata.Keys) {
        if ($FieldMap.ContainsKey($key)) {
            $converted[$FieldMap[$key]] = $Metadata[$key]
        }
        else {
            Add-StructureWarning "Sloupec '$key' v knihovně neexistuje, u '$FolderPath' ho přeskakuji."
        }
    }
    return $converted
}

function Set-MetadataColumns($List, $TargetColumns) {
    $fieldMap = Get-FieldNameMap $List

    foreach ($target in $TargetColumns) {
        # Existující sloupec hledáme podle interního i displejového názvu.
        $resolved = $null
        foreach ($alias in @($target.InternalName, $target.DisplayName)) {
            if ($alias -and $fieldMap.ContainsKey($alias)) { $resolved = $fieldMap[$alias]; break }
        }

        if ($resolved) {
            $note = if ($resolved -ne $target.InternalName) { " (interně $resolved)" } else { "" }
            Write-Host "  = $($target.InternalName)$note už existuje"
            continue
        }

        $internalName = Get-SafeInternalName $target.InternalName

        if ($Apply) {
            Add-PnPField -List $List.Id `
                -DisplayName $target.DisplayName `
                -InternalName $internalName `
                -Type Text `
                -AddToDefaultView | Out-Null
            Write-Host "  + $($target.DisplayName) vytvořen (interně $internalName, Text)" -ForegroundColor Green
        }
        else {
            Write-Host "  + [dry-run] vytvořil bych '$($target.DisplayName)' (interně $internalName, Text)" -ForegroundColor Yellow
        }
    }
}

Assert-Prerequisites

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

Write-Step "Čtu strukturu z $Path"
$rows = @(Import-StructureTable $Path $Delimiter)
if ($rows.Count -eq 0) { throw "Tabulka '$Path' je prázdná." }

$levelColumns = Get-LevelColumnNames $rows
Write-Host "  řádků: $($rows.Count)"
Write-Host "  úrovně zanoření: $($levelColumns -join ', ')"

$tableColumns = $rows[0].PSObject.Properties.Name
$unknownMeta = @($MetadataMap.Keys | Where-Object { $tableColumns -notcontains $_ })
if ($unknownMeta.Count -gt 0) {
    Add-StructureWarning "Sloupce z -MetadataMap, které v tabulce nejsou (ignoruji je): $($unknownMeta -join ', ')"
}

$desired = Get-DesiredFolders $rows $levelColumns $MetadataMap $FixedMetadata
$maxDepth = ($desired | Measure-Object -Property Depth -Maximum).Maximum
Write-Host "  složek k zajištění: $($desired.Count), nejhlubší úroveň: $maxDepth"

if ($FixedMetadata.Count -gt 0) {
    $fixedText = (($FixedMetadata.GetEnumerator() | Sort-Object Name |
                   ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; ")
    Write-Host "  konstantní metadata na každé složce: $fixedText"
}

Write-Step "Připojuji se k $SiteUrl"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$web = Get-PnPWeb
$list = Resolve-Library $Library
$libraryRoot = $list.RootFolder.ServerRelativeUrl
$librarySiteRelative = Get-SiteRelativeUrl $libraryRoot $web.ServerRelativeUrl

Write-Host "  knihovna: $($list.Title)"
Write-Host "  cesta:    $librarySiteRelative"
Write-Host "  Id:       $($list.Id)"

if ($list.BaseTemplate -ne 101) {
    Add-StructureWarning "'$($list.Title)' není knihovna dokumentů (BaseTemplate $($list.BaseTemplate)). Zkontrolujte parametr -Library."
}

Write-Step "Zjišťuji, co už v knihovně je"
$existing = Get-ExistingFolderMap $list $libraryRoot
Write-Host "  existujících složek: $($existing.Count)"

if (-not $SkipMetadata) {
    Write-Step "Sloupce pro metadata"
    Set-MetadataColumns $list (Get-TargetColumns $MetadataMap $FixedMetadata)
}

$plan = foreach ($folder in $desired) {
    [pscustomobject]@{
        Path      = $folder.Path
        Depth     = $folder.Depth
        Action    = if ($existing.ContainsKey($folder.Path)) { "skip" } else { "create" }
        Metadata  = (($folder.Metadata.GetEnumerator() | Sort-Object Name |
                      ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; ")
        FromTable = -not $folder.Implicit
    }
}

$toCreate = @($plan | Where-Object Action -eq "create")

Write-Step "Plán"
foreach ($item in $plan) {
    $indent = "  " * ($item.Depth - 1)
    $mark = if ($item.Action -eq "create") { "+" } else { "=" }
    $color = if ($item.Action -eq "create") { "Green" } else { "DarkGray" }
    $leafName = ($item.Path -split "/")[-1]
    $meta = if ($item.Metadata) { "   [$($item.Metadata)]" } else { "" }
    Write-Host "  $mark $indent$leafName$meta" -ForegroundColor $color
}

$plan | Export-Csv -Path "$OutputFolder/folder-plan.csv" -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "  vytvořit: $($toCreate.Count)     už existuje: $($plan.Count - $toCreate.Count)"
Write-Host "  plán uložen do $OutputFolder/folder-plan.csv"

function Write-WarningSummary {
    if ($script:Warnings.Count -gt 0) {
        $script:Warnings | Set-Content -Path "$OutputFolder/folder-warnings.txt" -Encoding UTF8
        Write-Host ""
        Write-Host "$($script:Warnings.Count) varování - podrobnosti v $OutputFolder/folder-warnings.txt" -ForegroundColor Yellow
    }
}

if (-not $Apply) {
    Write-Step "DRY-RUN - do SharePointu se nic nezapsalo"
    Write-Host "Až bude plán v pořádku, spusťte stejný příkaz s přepínačem -Apply." -ForegroundColor Yellow
    Write-WarningSummary
    return
}

Write-Step "Vytvářím složky"
$created = 0
foreach ($folder in ($desired | Where-Object { -not $existing.ContainsKey($_.Path) })) {
    $siteRelativePath = "$librarySiteRelative/$($folder.Path)"

    try {
        # Resolve-PnPFolder vytvoří celou cestu včetně chybějících nadřazených
        # složek a existující nechá být, takže je bezpečné ho pouštět opakovaně.
        Resolve-PnPFolder -SiteRelativePath $siteRelativePath -ErrorAction Stop | Out-Null
        $created++
        Write-Host "  + $($folder.Path)" -ForegroundColor Green
    }
    catch {
        Add-StructureWarning "Složku '$($folder.Path)' nelze vytvořit: $($_.Exception.Message)"
    }
}
Write-Host "  vytvořeno: $created"

if (-not $SkipMetadata) {
    Write-Step "Zapisuji metadata"

    # Znovu načíst - potřebujeme item ID právě vytvořených složek a interní
    # názvy právě založených sloupců.
    $existing = Get-ExistingFolderMap $list $libraryRoot
    $fieldMap = Get-FieldNameMap $list

    $updated = 0
    foreach ($folder in ($desired | Where-Object { $_.Metadata.Count -gt 0 })) {
        if (-not $existing.ContainsKey($folder.Path)) {
            Add-StructureWarning "Metadata pro '$($folder.Path)' nelze zapsat - složka v knihovně není."
            continue
        }

        $values = Convert-MetadataKeys $folder.Metadata $fieldMap $folder.Path
        if ($values.Count -eq 0) { continue }

        try {
            Set-PnPListItem -List $list.Id `
                -Identity $existing[$folder.Path] `
                -Values $values `
                -ErrorAction Stop | Out-Null
            $updated++
        }
        catch {
            Add-StructureWarning "Metadata pro '$($folder.Path)' nelze zapsat: $($_.Exception.Message)"
        }
    }
    Write-Host "  aktualizováno složek: $updated"
}

Write-WarningSummary
Write-Step "Hotovo"
