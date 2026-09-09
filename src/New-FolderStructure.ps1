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

.PARAMETER ListTerms
    Nic nevytváří. Pro zadaný sloupec se spravovanými metadaty vypíše termíny
    term setu, ke kterému je připojený, i s jejich GUIDy, a uloží je do
    export/terms-<sloupec>.csv. Slouží k dohledání přesného názvu termínu.

.PARAMETER UnlockReadOnlyFields
    Obyčejný sloupec označený ReadOnlyField zápis tiše zahodí. S tímto
    přepínačem ho skript před zápisem odemkne a po dokončení vrátí zpět na
    ReadOnly - i když zápis mezitím selže. Vyžaduje právo měnit sloupce knihovny.

    Pro sloupce se spravovanými metadaty tohle potřeba NENÍ - do těch se
    zapisuje přes CSOM, kterému příznak ReadOnly nevadí.

.PARAMETER ListFields
    Nic nevytváří. Vypíše sloupce knihovny s jejich interními názvy a typy
    a uloží je do export/library-fields.csv. Upozorní na sloupce, které mají
    stejný displejový název. Slouží k dohledání správného interního názvu.

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
    # Zjistit, jaké termíny sloupec se spravovanými metadaty přijímá
    ./src/New-FolderStructure.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Proj01" -Library "Dokumenty" -Path ./Folder_Structure.xlsx -ListTerms RevIMBCS

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

    # Jen vypíše sloupce knihovny a skončí. Slouží k dohledání interního názvu
    # a k odhalení několika sloupců se stejným displejovým názvem.
    [switch] $ListFields,

    # Vypíše termíny term setu, ke kterému je zadaný sloupec připojený, a skončí.
    # Slouží k dohledání přesného názvu nebo GUIDu termínu.
    [string] $ListTerms = "",

    # Sloupce označené ReadOnlyField na dobu zápisu odemkne a na konci je vrátí
    # zpět na ReadOnly, i když zápis selže.
    [switch] $UnlockReadOnlyFields,

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

function Write-WarningSummary {
    if ($script:Warnings.Count -gt 0) {
        $script:Warnings | Set-Content -Path "$OutputFolder/folder-warnings.txt" -Encoding UTF8
        Write-Host ""
        Write-Host "$($script:Warnings.Count) varování - podrobnosti v $OutputFolder/folder-warnings.txt" -ForegroundColor Yellow
    }
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
    $allLists = Get-PnPList -Includes RootFolder, ForceCheckout

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
        # Držíme celý objekt položky, ne jen Id - zápis do sloupce se
        # spravovanými metadaty jde přes CSOM a potřebuje ListItem.
        if ($relative) { $map[$relative] = $item }
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

# Index sloupců knihovny. Jeden název může ukazovat na víc sloupců - knihovna
# běžně obsahuje několik sloupců se stejným displejovým názvem, ale různým
# interním názvem a typem. Proto se pod každým aliasem drží SEZNAM, ne jeden
# sloupec: vybrat naslepo první by znamenalo zapisovat do cizího sloupce.
function Get-FieldIndex($List) {
    $index = @{}

    foreach ($field in (Get-PnPField -List $List.Id)) {
        $descriptor = [pscustomobject]@{
            InternalName = $field.InternalName
            Title        = $field.Title
            Type         = $field.TypeAsString
            Hidden       = $field.Hidden
            ReadOnly     = $field.ReadOnlyField
            Sealed       = $field.Sealed
        }

        foreach ($alias in @($field.InternalName, $field.StaticName, $field.Title)) {
            if (-not $alias) { continue }
            if (-not $index.ContainsKey($alias)) { $index[$alias] = @() }
            if ($index[$alias].InternalName -notcontains $field.InternalName) {
                $index[$alias] += $descriptor
            }
        }
    }

    return $index
}

# Dohledá sloupec podle zadaného názvu. Přesná shoda s interním názvem má
# přednost před displejovým - jinak by se nešlo z nejednoznačnosti dostat.
# Když je názvů víc a interní se netrefí, vrací $null a vypíše kandidáty.
function Resolve-Field($Name, $FieldIndex) {
    if (-not $FieldIndex.ContainsKey($Name)) { return $null }

    $candidates = @($FieldIndex[$Name])
    if ($candidates.Count -eq 1) { return $candidates[0] }

    $exact = @($candidates | Where-Object { $_.InternalName -eq $Name })
    if ($exact.Count -eq 1) { return $exact[0] }

    $list = ($candidates | ForEach-Object { "$($_.InternalName) ($($_.Type))" }) -join ", "
    Add-StructureWarning "Název '$Name' odpovídá $($candidates.Count) sloupcům: $list. Zadejte místo něj interní název toho správného."
    return $null
}

function Test-IsTaxonomyField($Field) {
    return $Field.Type -eq "TaxonomyFieldType" -or $Field.Type -eq "TaxonomyFieldTypeMulti"
}

# Interní název pro nově zakládaný sloupec. Mezery a diakritika by se zakódovaly
# do nečitelného _x0020_, takže je rovnou vynecháme a hezký název dáme do titulku.
function Get-SafeInternalName($Name) {
    $clean = ($Name -replace '[^A-Za-z0-9_]', '')
    if (-not $clean) { throw "Z názvu sloupce '$Name' nelze odvodit interní název." }
    if ($clean -match '^\d') { $clean = "f$clean" }
    return $clean
}

# Rozdělí metadata na dvě skupiny, protože se zapisují jinak:
#   Plain    - obyčejné sloupce, jde na ně Set-PnPListItem
#   Taxonomy - spravovaná metadata, jdou přes CSOM (i na ReadOnly poli)
function Convert-MetadataKeys($Metadata, $FieldIndex, $FolderPath, $List) {
    $plain = @{}
    $taxonomy = @()

    foreach ($key in $Metadata.Keys) {
        $field = Resolve-Field $key $FieldIndex

        if (-not $field) {
            Add-StructureWarning "Sloupec '$key' nelze u '$FolderPath' jednoznačně určit, přeskakuji ho."
            continue
        }

        if ($field.Sealed) {
            Add-StructureWarning "Sloupec '$key' ($($field.InternalName)) je Sealed - mění se jen na úrovni content typu, ne tady."
        }

        if (Test-IsTaxonomyField $field) {
            $term = Resolve-Term $List $field.InternalName $Metadata[$key]

            if (-not $term) {
                Add-StructureWarning "Metadata '$key' u '$FolderPath' přeskakuji - termín '$($Metadata[$key])' se nepodařilo přeložit."
                continue
            }

            $taxonomy += [pscustomobject]@{
                InternalName = $field.InternalName
                Term         = $term
            }
            continue
        }

        # Do obyčejného sloupce označeného ReadOnlyField SharePoint zápis přes
        # Set-PnPListItem tiše zahodí - projde bez chyby, hodnota se neuloží.
        if ($field.ReadOnly -and -not $UnlockReadOnlyFields) {
            Add-StructureWarning "Sloupec '$key' ($($field.InternalName)) je ReadOnly, zápis by se zahodil. Použijte -UnlockReadOnlyFields, nebo ho odemkněte ručně."
            continue
        }

        $plain[$field.InternalName] = $Metadata[$key]
    }

    return [pscustomobject]@{ Plain = $plain; Taxonomy = @($taxonomy) }
}

function Set-MetadataColumns($List, $TargetColumns) {
    $fieldIndex = Get-FieldIndex $List

    foreach ($target in $TargetColumns) {
        # Existující sloupec hledáme podle interního i displejového názvu.
        $field = $null
        foreach ($alias in @($target.InternalName, $target.DisplayName)) {
            if (-not $alias) { continue }
            if ($fieldIndex.ContainsKey($alias)) {
                $field = Resolve-Field $alias $fieldIndex
                break
            }
        }

        if ($field) {
            $note = if ($field.InternalName -ne $target.InternalName) { " -> $($field.InternalName)" } else { "" }
            $flags = @()
            if ($field.ReadOnly) { $flags += "ReadOnly" }
            if ($field.Sealed)   { $flags += "Sealed" }
            if ($field.Hidden)   { $flags += "Hidden" }
            $flagText = if ($flags.Count -gt 0) { "   [$($flags -join ', ')]" } else { "" }

            $problem = (Test-IsTaxonomyField $field) -or $field.ReadOnly -or $field.Sealed
            $color = if ($problem) { "Yellow" } else { "Gray" }
            Write-Host "  = $($target.InternalName)$note   typ $($field.Type)$flagText" -ForegroundColor $color

            continue
        }

        # Sloupec, jehož název byl nejednoznačný, se nezakládá - jeden takový
        # už existuje a další duplikát by problém jen zhoršil.
        if ($fieldIndex.ContainsKey($target.InternalName) -or $fieldIndex.ContainsKey($target.DisplayName)) {
            Write-Host "  ! $($target.InternalName) nelze určit, nový nezakládám" -ForegroundColor Red
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

# Přepne ReadOnlyField na $false a vrátí seznam sloupců, které přepnula, aby
# se daly vrátit zpět. Volající to MUSÍ vrátit ve finally bloku.
function Disable-FieldReadOnly($List, $FieldNames) {
    $unlocked = @()

    foreach ($name in $FieldNames) {
        try {
            Set-PnPField -List $List.Id -Identity $name -Values @{ ReadOnlyField = $false } -ErrorAction Stop | Out-Null
            $unlocked += $name
            Write-Host "  odemčen $name" -ForegroundColor Yellow
        }
        catch {
            Add-StructureWarning "Sloupec '$name' nelze odemknout: $($_.Exception.Message)"
        }
    }

    return @($unlocked)
}

function Enable-FieldReadOnly($List, $FieldNames) {
    foreach ($name in $FieldNames) {
        try {
            Set-PnPField -List $List.Id -Identity $name -Values @{ ReadOnlyField = $true } -ErrorAction Stop | Out-Null
            Write-Host "  zamčen zpět $name" -ForegroundColor Yellow
        }
        catch {
            # Tohle je nutné říct nahlas - sloupec zůstal zapisovatelný.
            Add-StructureWarning "POZOR: sloupec '$name' se nepodařilo vrátit na ReadOnly: $($_.Exception.Message). Vraťte ho ručně."
        }
    }
}

$script:TermCache = @{}

# Sloupec se spravovanými metadaty jako CSOM objekt. Načítá se přes kolekci
# Fields, protože CSOM vrátí rovnou typ TaxonomyField - ten má TermSetId
# i metodu SetFieldValueByValue, které obyčejný Field nemá.
function Get-TaxonomyField($List, $InternalName) {
    try {
        $fields = Get-PnPProperty -ClientObject $List -Property Fields -ErrorAction Stop
    }
    catch {
        Add-StructureWarning "Sloupce knihovny nelze načíst: $($_.Exception.Message)"
        return $null
    }

    $field = $fields | Where-Object { $_.InternalName -eq $InternalName } | Select-Object -First 1
    if (-not $field) {
        Add-StructureWarning "Sloupec '$InternalName' v knihovně není."
        return $null
    }

    return $field
}

# Termíny term setu, ke kterému je sloupec připojený. Jde se přímo z pole na
# jeho TermSetId, takže není potřeba hledat skupinu v Term Store.
# Postup převzatý ze Set-CsdClass.ps1 (autor Sergiu Nica).
function Get-TermsForField($List, $InternalName) {
    if ($script:TermCache.ContainsKey($InternalName)) {
        return $script:TermCache[$InternalName]
    }

    $result = [pscustomobject]@{ Field = $null; Terms = @() }

    $field = Get-TaxonomyField $List $InternalName
    if (-not $field) {
        $script:TermCache[$InternalName] = $result
        return $result
    }
    $result.Field = $field

    if (-not $field.TermSetId) {
        Add-StructureWarning "U sloupce '$InternalName' se nepodařilo zjistit TermSetId. Je to opravdu sloupec se spravovanými metadaty?"
        $script:TermCache[$InternalName] = $result
        return $result
    }

    try {
        $context = Get-PnPContext
        $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
        $termStore = $session.GetDefaultSiteCollectionTermStore()
        $termSet = $termStore.GetTermSet([Guid]$field.TermSetId)
        $terms = $termSet.GetAllTerms()

        $context.Load($terms)
        $context.ExecuteQuery()

        $result.Terms = @($terms)
    }
    catch {
        Add-StructureWarning "Termíny pro '$InternalName' nelze načíst: $($_.Exception.Message). Máte přístup na Term Store?"
    }

    $script:TermCache[$InternalName] = $result
    return $result
}

$script:TermResolution = @{}

# Termíny v klasifikačních schématech začínají číslem ("5.3 Car Series and
# Concept Docs"). Číslo je stabilní, text za ním se v Term Store liší
# formulací, pomlčkou nebo velikostí písmen - proto se porovnává hlavně ono.
function Get-TermNumber($Name) {
    if ("$Name" -match '^\s*(\d+(?:\.\d+)*)') { return $Matches[1] }
    return $null
}

function Get-NormalizedTermName($Name) {
    return (("$Name" -replace '\s+', ' ').Trim())
}

# Termín podle názvu, čísla, nebo GUIDu. Vrací objekt termínu, nebo $null.
# Výsledek se pamatuje, aby stejný termín nehlásil chybu u každé složky zvlášť.
function Resolve-Term($List, $InternalName, $Label) {
    $cacheKey = "$InternalName|$Label"
    if ($script:TermResolution.ContainsKey($cacheKey)) {
        return $script:TermResolution[$cacheKey]
    }

    $script:TermResolution[$cacheKey] = $null

    $source = Get-TermsForField $List $InternalName
    if ($source.Terms.Count -eq 0) { return $null }

    $wanted = Get-NormalizedTermName $Label
    $wantedNumber = Get-TermNumber $wanted

    # 1. GUID
    $match = @($source.Terms | Where-Object { $_.Id.ToString() -eq $wanted })

    # 2. přesný název
    if ($match.Count -eq 0) {
        $match = @($source.Terms | Where-Object { (Get-NormalizedTermName $_.Name) -eq $wanted })
    }

    # 3. shoda čísla na začátku - "5.3" i "5.3 Cokoliv" najde termín číslo 5.3
    if ($match.Count -eq 0 -and $wantedNumber) {
        $match = @($source.Terms | Where-Object { (Get-TermNumber $_.Name) -eq $wantedNumber })
    }

    # 4. jeden název je začátkem druhého
    if ($match.Count -eq 0) {
        $match = @($source.Terms | Where-Object {
            $name = Get-NormalizedTermName $_.Name
            $name.StartsWith($wanted, [System.StringComparison]::OrdinalIgnoreCase) -or
            $wanted.StartsWith($name, [System.StringComparison]::OrdinalIgnoreCase)
        })
    }

    if ($match.Count -eq 1) {
        $found = $match[0]
        if ((Get-NormalizedTermName $found.Name) -ne $wanted) {
            Write-Host "  termín '$Label' -> '$($found.Name)'" -ForegroundColor DarkGray
        }
        $script:TermResolution[$cacheKey] = $found
        return $found
    }

    if ($match.Count -gt 1) {
        $names = ($match | Select-Object -First 5 | ForEach-Object { "'$($_.Name)'" }) -join ", "
        Add-StructureWarning "Termín '$Label' odpovídá víc termínům: $names. Předejte GUID toho správného."
        return $null
    }

    # Nenašlo se - vypsat, co term set obsahuje, ať se to nemusí hledat jinde.
    Add-StructureWarning "Termín '$Label' v term setu není."
    Show-AvailableTerms $source.Terms $InternalName
    return $null
}

function Show-AvailableTerms($Terms, $InternalName) {
    Write-Host ""
    Write-Host "  Term set sloupce $InternalName obsahuje $($Terms.Count) termínů:" -ForegroundColor Yellow

    $sorted = @($Terms | Sort-Object Name)
    foreach ($term in ($sorted | Select-Object -First 30)) {
        Write-Host "    $($term.Name)" -ForegroundColor DarkGray
    }
    if ($sorted.Count -gt 30) {
        Write-Host "    ... a dalších $($sorted.Count - 30)" -ForegroundColor DarkGray
    }

    try {
        $target = "$OutputFolder/terms-$InternalName.csv"
        $sorted | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Id = $_.Id.ToString() } } |
            Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
        Write-Host "  Úplný seznam v $target" -ForegroundColor Yellow
    }
    catch {
        Write-Verbose "Seznam termínů nelze uložit: $($_.Exception.Message)"
    }
    Write-Host ""
}

# Zápis do sloupce se spravovanými metadaty přes CSOM. Set-PnPListItem tady
# nepomůže - u pole označeného ReadOnly hodnotu zahodí. SetFieldValueByValue
# jde pod tím a projde. Postup převzatý ze Set-CsdClass.ps1.
# Volající musí po všech zápisech zavolat Invoke-PnPQuery.
function Set-TaxonomyFieldValue($Field, $Item, $Term) {
    $value = New-Object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue
    $value.Label = $Term.Name
    $value.TermGuid = $Term.Id.ToString()
    # -1 znamená, že si SharePoint dohledá WssId sám.
    $value.WssId = -1

    $Field.SetFieldValueByValue($Item, $value)
    $Item.Update()
}

function Show-TermSet($List, $FieldInternalName, $OutFolder) {
    $source = Get-TermsForField $List $FieldInternalName

    if ($source.Field) {
        Write-Host "  sloupec:   $($source.Field.InternalName)  ($($source.Field.Title))"
        Write-Host "  typ:       $($source.Field.TypeAsString)"
        Write-Host "  TermSetId: $($source.Field.TermSetId)"
    }
    Write-Host ""

    if ($source.Terms.Count -eq 0) {
        Write-Host "  Žádné termíny se nenačetly." -ForegroundColor Yellow
        return
    }

    $rows = $source.Terms | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Id = $_.Id.ToString() }
    }

    $rows | Sort-Object Name | Format-Table Name, Id -AutoSize | Out-String -Width 220 | Write-Host

    $target = "$OutFolder/terms-$FieldInternalName.csv"
    $rows | Sort-Object Name | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    Write-Host "  $($rows.Count) termínů, seznam v $target"
}

function Show-LibraryFields($List, $OutFolder) {
    $all = Get-PnPField -List $List.Id |
        Select-Object Title, InternalName, StaticName, TypeAsString, ReadOnlyField, Sealed,
                      Hidden, Required, FromBaseType, Group |
        Sort-Object Title, InternalName

    # Vestavěné sloupce (Author, Created, ID, ...) a interní sloupce s podtržítkem
    # jsou read-only vždycky a nikoho nezajímají. Zajímavé jsou ty, které do
    # knihovny přidal někdo nebo nějaké řešení - do těch se dá chtít zapisovat.
    $custom = @($all | Where-Object {
        -not $_.FromBaseType -and -not $_.InternalName.StartsWith("_")
    })

    Write-Host "  Sloupce přidané do knihovny ($($custom.Count) z $($all.Count) celkem):"
    Write-Host ""
    $custom | Format-Table Title, InternalName, TypeAsString, ReadOnlyField, Sealed, Hidden, Group -AutoSize |
        Out-String -Width 220 | Write-Host

    $writable = @($custom | Where-Object { -not $_.ReadOnlyField -and -not $_.Sealed -and -not $_.Hidden })
    Write-Host "  Zapisovatelné: $(($writable.InternalName | Sort-Object) -join ', ')"
    Write-Host ""

    $blocked = @($custom | Where-Object { $_.ReadOnlyField -or $_.Sealed })
    foreach ($field in $blocked) {
        $why = @()
        if ($field.ReadOnlyField) { $why += "ReadOnly" }
        if ($field.Sealed) { $why += "Sealed" }
        Write-Host "  ZAMČENO: '$($field.Title)' ($($field.InternalName)), typ $($field.TypeAsString): $($why -join ', ')" -ForegroundColor Yellow
    }

    # Duplicitní displejové názvy - kvůli nim nelze sloupec určit podle názvu.
    $duplicates = $all | Group-Object Title | Where-Object { $_.Count -gt 1 }
    foreach ($group in $duplicates) {
        $names = ($group.Group | ForEach-Object { "$($_.InternalName) ($($_.TypeAsString))" }) -join ", "
        Write-Host "  DUPLICITA: název '$($group.Name)' má $($group.Count) sloupců: $names" -ForegroundColor Red
    }

    $target = "$OutFolder/library-fields.csv"
    $all | Export-Csv -Path $target -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Úplný seznam včetně vestavěných sloupců v $target"
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

# Uzamčení celé webové kolekce se z tohoto skriptu odemknout nedá - je na to
# potřeba SharePoint Administrator a připojení do admin centra. Aspoň to ale
# poznáme a řekneme, místo aby zápisy tiše nic nedělaly.
try {
    $site = Get-PnPSite -Includes ReadOnly -ErrorAction Stop
    if ($site.ReadOnly) {
        Add-StructureWarning "Celá webová kolekce je v režimu ReadOnly - žádný zápis neprojde. Odemčení vyžaduje SharePoint Administrator, viz docs/07-struktura-slozek-z-excelu.md."
    }
}
catch {
    Write-Verbose "Stav uzamčení webu nelze zjistit: $($_.Exception.Message)"
}

if ($list.ForceCheckout) {
    Add-StructureWarning "Knihovna vyžaduje Check-out (ForceCheckout). Zápis metadat na složky tím může být blokovaný."
}

Write-Step "Zjišťuji, co už v knihovně je"
$existing = Get-ExistingFolderMap $list $libraryRoot
Write-Host "  existujících složek: $($existing.Count)"

if ($ListFields) {
    Write-Step "Sloupce v knihovně $($list.Title)"
    Show-LibraryFields $list $OutputFolder
    return
}

if ($ListTerms) {
    Write-Step "Termíny pro sloupec $ListTerms"
    Show-TermSet $list $ListTerms $OutputFolder
    Write-WarningSummary
    return
}

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
    $fieldIndex = Get-FieldIndex $list

    # Sloupce, do kterých se má zapisovat a které jsou zamčené na ReadOnly.
    $lockedFields = @()
    if ($UnlockReadOnlyFields) {
        $wantedKeys = @($desired | ForEach-Object { $_.Metadata.Keys } | Sort-Object -Unique)
        foreach ($key in $wantedKeys) {
            $field = Resolve-Field $key $fieldIndex
            if ($field -and $field.ReadOnly) { $lockedFields += $field.InternalName }
        }
        $lockedFields = @($lockedFields | Sort-Object -Unique)
    }

    $unlockedFields = @()

    try {
        if ($lockedFields.Count -gt 0) {
            Write-Host "  odemykám $($lockedFields.Count) ReadOnly sloupců na dobu zápisu" -ForegroundColor Yellow
            $unlockedFields = Disable-FieldReadOnly $list $lockedFields

            # Po odemčení je index zastaralý - příznak ReadOnly už neplatí.
            $fieldIndex = Get-FieldIndex $list
        }

        $updated = 0
        $taxonomyWrites = 0

        foreach ($folder in ($desired | Where-Object { $_.Metadata.Count -gt 0 })) {
            if (-not $existing.ContainsKey($folder.Path)) {
                Add-StructureWarning "Metadata pro '$($folder.Path)' nelze zapsat - složka v knihovně není."
                continue
            }

            $item = $existing[$folder.Path]
            $values = Convert-MetadataKeys $folder.Metadata $fieldIndex $folder.Path $list
            $touched = $false

            if ($values.Plain.Count -gt 0) {
                try {
                    Set-PnPListItem -List $list.Id `
                        -Identity $item.Id `
                        -Values $values.Plain `
                        -ErrorAction Stop | Out-Null
                    $touched = $true
                }
                catch {
                    Add-StructureWarning "Metadata pro '$($folder.Path)' nelze zapsat: $($_.Exception.Message)"
                }
            }

            # Spravovaná metadata se zapisují přes CSOM a odešlou se dávkou
            # v Invoke-PnPQuery, až projdou všechny složky.
            foreach ($assignment in $values.Taxonomy) {
                $source = Get-TermsForField $list $assignment.InternalName
                if (-not $source.Field) { continue }

                try {
                    Set-TaxonomyFieldValue $source.Field $item $assignment.Term
                    $taxonomyWrites++
                    $touched = $true
                }
                catch {
                    Add-StructureWarning "Termín do '$($assignment.InternalName)' u '$($folder.Path)' nelze nastavit: $($_.Exception.Message)"
                }
            }

            if ($touched) { $updated++ }
        }

        if ($taxonomyWrites -gt 0) {
            Write-Host "  odesílám $taxonomyWrites zápisů spravovaných metadat"
            try {
                Invoke-PnPQuery
            }
            catch {
                Add-StructureWarning "Zápis spravovaných metadat se nepodařilo odeslat: $($_.Exception.Message)"
            }
        }

        Write-Host "  aktualizováno složek: $updated"
    }
    finally {
        # Zamknout zpět za všech okolností, i když zápis spadl nebo ho někdo
        # přerušil - jinak by knihovna zůstala otevřená k editaci.
        if ($unlockedFields.Count -gt 0) {
            Write-Host ""
            Write-Host "  vracím sloupce na ReadOnly" -ForegroundColor Yellow
            Enable-FieldReadOnly $list $unlockedFields
        }
    }
}

Write-WarningSummary
Write-Step "Hotovo"
