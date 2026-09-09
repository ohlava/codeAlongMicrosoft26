<#
.SYNOPSIS
    Nastaví CSD Class (nebo jiný sloupec) v knihovně - jako výchozí hodnotu pro
    nově nahrávané soubory a volitelně i na soubory, které v ní už jsou.

.DESCRIPTION
    Vychází ze Set-CsdClass.ps1 (autor Sergiu Nica). Klíčová část je zápis do
    sloupce se spravovanými metadaty, který je označený ReadOnly:
    Set-PnPListItem by hodnotu zahodil, ale CSOM SetFieldValueByValue projde.

    Dělá dvě věci, které se dají volit přes -Scope:

      DefaultValue   výchozí hodnota sloupce v knihovně - dostane ji každý
                     NOVĚ nahraný soubor
      ExistingFiles  označí soubory, které už v knihovně jsou

    VÝCHOZÍ REŽIM JE DRY-RUN. Bez -Apply se nic nezapíše.

.PARAMETER SiteUrl
    Web s knihovnou.

.PARAMETER Library
    Knihovna dokumentů. GUID, název, nebo cesta relativní k webu.

.PARAMETER Field
    Interní název sloupce, např. RevIMBCS.

.PARAMETER Value
    Hodnota. U spravovaných metadat název termínu, jeho číslo ("5.3"), nebo GUID.

.PARAMETER Title
    Popisek, pod kterým se sloupec zobrazuje v knihovně. Například "CSD Class"
    místo výchozího "Třída KSU". Mění se jen v této knihovně, ne globálně
    v Term Store ani u jiných webů.

.PARAMETER Scope
    DefaultValue, ExistingFiles, nebo Both. Výchozí DefaultValue.

.PARAMETER ClientId
    Client ID Entra ID aplikace. Když se nezadá, vezme se z PNP_CLIENT_ID.

.PARAMETER Apply
    Provede změny.

.EXAMPLE
    # Co by se stalo
    ./src/Set-CsdClass.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/P42" -Library "Dokumenty" -Field RevIMBCS -Value "5.3 Car Series and Concept Docs"

.EXAMPLE
    # Výchozí hodnota pro nové soubory i označení existujících
    ./src/Set-CsdClass.ps1 -SiteUrl "..." -Library "Dokumenty" -Field RevIMBCS -Value "5.3 Car Series and Concept Docs" -Scope Both -Apply

.EXAMPLE
    # Vypsat dostupné termíny
    ./src/Set-CsdClass.ps1 -SiteUrl "..." -Library "Dokumenty" -Field RevIMBCS -ListTerms

.NOTES
    Postup: docs/12-nastaveni-a-spusteni.md
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $Library,
    [Parameter(Mandatory = $true)] [string] $Field,

    [string] $Value = "",

    # Popisek sloupce v knihovně. Změní se jen v této knihovně, ne globálně.
    [string] $Title = "",

    [ValidateSet("DefaultValue", "ExistingFiles", "Both")]
    [string] $Scope = "DefaultValue",

    [string] $ClientId = $env:PNP_CLIENT_ID,

    [switch] $ListTerms,
    [switch] $Apply
)

$ErrorActionPreference = "Stop"

function Write-Step($Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "Modul PnP.PowerShell není nainstalovaný. Spusťte: Install-Module PnP.PowerShell -Scope CurrentUser"
}
if (-not $ClientId) {
    throw "Chybí ClientId. Zadejte -ClientId nebo nastavte PNP_CLIENT_ID."
}
if (-not $ListTerms -and -not $Value) {
    throw "Chybí -Value. Nebo použijte -ListTerms a podívejte se, co sloupec přijímá."
}

Write-Step "Připojuji se k $SiteUrl"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ============================================================
# Knihovna a sloupec
# ============================================================

$allLists = Get-PnPList -Includes RootFolder
$list = $allLists | Where-Object { $_.Id.ToString() -eq $Library } | Select-Object -First 1
if (-not $list) { $list = $allLists | Where-Object { $_.Title -eq $Library } | Select-Object -First 1 }
if (-not $list) {
    $list = $allLists | Where-Object {
        $_.RootFolder.ServerRelativeUrl.EndsWith("/$Library")
    } | Select-Object -First 1
}
if (-not $list) {
    throw "Knihovnu '$Library' jsem nenašel."
}

# Fields přes CSOM, protože u sloupce se spravovanými metadaty tak dostaneme
# rovnou typ TaxonomyField - ten má TermSetId i SetFieldValueByValue.
$fields = Get-PnPProperty -ClientObject $list -Property Fields
$target = $fields | Where-Object { $_.InternalName -eq $Field } | Select-Object -First 1

if (-not $target) {
    $candidates = ($fields | Where-Object { -not $_.Hidden -and -not $_.FromBaseType } |
        ForEach-Object { "  $($_.InternalName)  ($($_.Title), $($_.TypeAsString))" }) -join "`n"
    throw "Sloupec '$Field' v knihovně '$($list.Title)' není. Dostupné:`n$candidates"
}

$isTaxonomy = $target.TypeAsString -eq "TaxonomyFieldType" -or $target.TypeAsString -eq "TaxonomyFieldTypeMulti"

Write-Host "  knihovna: $($list.Title)"
Write-Host "  sloupec:  $($target.InternalName)  ($($target.Title))"
Write-Host "  typ:      $($target.TypeAsString)$(if ($target.ReadOnlyField) { '   ReadOnly' })"

# ============================================================
# Termín
# ============================================================

$term = $null

if ($isTaxonomy) {
    Write-Step "Načítám termíny z Term Store"

    $context = Get-PnPContext
    $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
    $termStore = $session.GetDefaultSiteCollectionTermStore()
    $termSet = $termStore.GetTermSet([Guid]$target.TermSetId)
    $terms = $termSet.GetAllTerms()
    $context.Load($terms)
    $context.ExecuteQuery()

    Write-Host "  term set: $($target.TermSetId)"
    Write-Host "  termínů:  $($terms.Count)"

    if ($ListTerms) {
        Write-Step "Dostupné termíny"
        $terms | Sort-Object Name | ForEach-Object {
            Write-Host ("  {0,-45} {1}" -f $_.Name, $_.Id)
        }
        return
    }

    # Termíny klasifikačního schématu začínají číslem ("5.3 ..."). Číslo je
    # stabilní, text za ním se v Term Store liší formulací nebo velikostí
    # písmen, takže se porovnává hlavně ono.
    function Get-TermNumber($Name) {
        if ("$Name" -match '^\s*(\d+(?:\.\d+)*)') { return $Matches[1] }
        return $null
    }
    function Get-NormalizedName($Name) { return (("$Name" -replace '\s+', ' ').Trim()) }

    $wanted = Get-NormalizedName $Value
    $wantedNumber = Get-TermNumber $wanted

    $candidates = @($terms | Where-Object { $_.Id.ToString() -eq $wanted })

    if ($candidates.Count -eq 0) {
        $candidates = @($terms | Where-Object { (Get-NormalizedName $_.Name) -eq $wanted })
    }
    if ($candidates.Count -eq 0 -and $wantedNumber) {
        $candidates = @($terms | Where-Object { (Get-TermNumber $_.Name) -eq $wantedNumber })
    }
    if ($candidates.Count -eq 0) {
        $candidates = @($terms | Where-Object {
            $name = Get-NormalizedName $_.Name
            $name.StartsWith($wanted, [System.StringComparison]::OrdinalIgnoreCase) -or
            $wanted.StartsWith($name, [System.StringComparison]::OrdinalIgnoreCase)
        })
    }

    if ($candidates.Count -gt 1) {
        $names = ($candidates | ForEach-Object { "  $($_.Name)" }) -join "`n"
        throw "Hodnota '$Value' odpovídá víc termínům:`n$names`nPředejte GUID toho správného."
    }

    if ($candidates.Count -eq 0) {
        $all = ($terms | Sort-Object Name | Select-Object -First 30 |
            ForEach-Object { "  $($_.Name)" }) -join "`n"
        $more = if ($terms.Count -gt 30) { "`n  ... a dalších $($terms.Count - 30)" } else { "" }
        throw "Termín '$Value' v term setu není.`n`nTerm set obsahuje $($terms.Count) termínů:`n$all$more"
    }

    $term = $candidates[0]

    if ((Get-NormalizedName $term.Name) -ne $wanted) {
        Write-Host "  zadáno:   $Value" -ForegroundColor DarkGray
        Write-Host "  nalezeno: $($term.Name)   (shoda podle čísla)" -ForegroundColor Yellow
    }

    Write-Host "  termín:   $($term.Name)"
    Write-Host "  GUID:     $($term.Id)"
}
elseif ($ListTerms) {
    Write-Host ""
    Write-Host "  Sloupec není typu spravovaná metadata, žádné termíny nemá." -ForegroundColor Yellow
    return
}

# ============================================================
# Popisek sloupce
# ============================================================

if ($Title -and $target.Title -ne $Title) {
    Write-Step "Popisek sloupce"
    Write-Host "  '$($target.Title)' -> '$Title'"

    if ($Apply) {
        try {
            Set-PnPField -List $list.Id -Identity $target.InternalName -Values @{ Title = $Title } -ErrorAction Stop
            Write-Host "  přejmenováno" -ForegroundColor Green
        }
        catch {
            Write-Host "  nelze přejmenovat: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "  [náhled] přejmenoval bych" -ForegroundColor Yellow
    }
}

if (-not $Apply) {
    Write-Step "DRY-RUN - nic se nezapsalo"

    if ($Scope -ne "ExistingFiles") {
        Write-Host "  nastavil bych výchozí hodnotu sloupce v knihovně" -ForegroundColor Yellow
    }
    if ($Scope -ne "DefaultValue") {
        $files = Get-PnPListItem -List $list.Id -PageSize 2000 -Fields "FileLeafRef", "FSObjType"
        $fileCount = @($files | Where-Object { $_.FieldValues.FSObjType -eq 0 }).Count
        Write-Host "  označil bych $fileCount existujících souborů" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Až bude výpis v pořádku, spusťte stejný příkaz s -Apply." -ForegroundColor Yellow
    return
}

# ============================================================
# Výchozí hodnota sloupce
# ============================================================

if ($Scope -ne "ExistingFiles") {
    Write-Step "Nastavuji výchozí hodnotu sloupce"

    # U spravovaných metadat má výchozí hodnota tvar "-1;#Nazev|GUID".
    $defaultValue = if ($isTaxonomy) { "-1;#$($term.Name)|$($term.Id)" } else { $Value }

    # Cmdlet se mezi verzemi PnP jmenuje jednou v jednotném, jednou v množném
    # čísle. Vybereme ten, který v nainstalované verzi opravdu je.
    $cmdlet = @("Set-PnPDefaultColumnValues", "Set-PnPDefaultColumnValue") |
        Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } |
        Select-Object -First 1

    if (-not $cmdlet) {
        Write-Host "  přeskakuji: nainstalovaná verze PnP.PowerShell nezná Set-PnPDefaultColumnValues." -ForegroundColor Yellow
        Write-Host "  Nastavte výchozí hodnotu ručně: knihovna -> Nastavení -> Výchozí hodnoty sloupců." -ForegroundColor DarkGray
    }
    else {
        try {
            & $cmdlet -List $list.Id -Field $target.InternalName -Value $defaultValue -ErrorAction Stop
            Write-Host "  nastaveno: $defaultValue" -ForegroundColor Green
            Write-Host "  Platí pro nově nahrávané soubory, existující nemění." -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  nelze nastavit ($cmdlet): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ============================================================
# Existující soubory
# ============================================================

if ($Scope -ne "DefaultValue") {
    Write-Step "Označuji existující soubory"

    $items = Get-PnPListItem -List $list.Id -PageSize 500 -Fields "FileLeafRef", "FSObjType", $target.InternalName
    $files = @($items | Where-Object { $_.FieldValues.FSObjType -eq 0 })

    Write-Host "  souborů: $($files.Count)"

    $updated = 0
    $failed = 0

    foreach ($item in $files) {
        try {
            if ($isTaxonomy) {
                # Set-PnPListItem by u ReadOnly pole hodnotu zahodil, CSOM projde.
                $taxonomyValue = New-Object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue
                $taxonomyValue.Label = $term.Name
                $taxonomyValue.TermGuid = $term.Id.ToString()
                $taxonomyValue.WssId = -1

                $target.SetFieldValueByValue($item, $taxonomyValue)
            }
            else {
                $item[$target.InternalName] = $Value
            }

            $item.Update()
            $updated++
        }
        catch {
            $failed++
            Write-Host "  ! $($item.FieldValues.FileLeafRef): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($updated -gt 0) {
        Write-Host "  odesílám $updated změn"
        try {
            Invoke-PnPQuery
            Write-Host "  zapsáno" -ForegroundColor Green
        }
        catch {
            Write-Host "  zápis se nepodařilo odeslat: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "  označeno: $updated, chyb: $failed"
}

Write-Step "Hotovo"
