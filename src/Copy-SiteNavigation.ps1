<#
.SYNOPSIS
    Zkopíruje navigaci z jednoho SharePointového webu na druhý.

.DESCRIPTION
    Přenese levou navigaci (QuickLaunch), horní navigaci (TopNavigationBar),
    nebo obojí, v libovolné hloubce zanoření. Odkazy míříci do zdrojového webu
    se přepíšou na cílový web, odkazy mimo něj zůstanou nedotčené.

    VÝCHOZÍ REŽIM JE DRY-RUN. Bez přepínače -Apply se nic nezmění, jen se vypíše
    a uloží plán.

    Výchozí režim slučování (-Mode Merge) do cílové navigace jen doplní chybějící
    položky, existující nechá být, takže se dá pouštět opakovaně. Režim -Mode
    Replace nejdřív smaže celou cílovou navigaci - to je nevratné a je potřeba
    ho vyžádat výslovně.

.PARAMETER SourceSiteUrl
    Web, ze kterého se navigace čte. Jen se z něj čte.

.PARAMETER TargetSiteUrl
    Web, na který se navigace přenáší.

.PARAMETER ClientId
    Client ID Entra ID aplikace. Když se nezadá, vezme se z PNP_CLIENT_ID.

.PARAMETER Location
    QuickLaunch (levá navigace), TopNavigationBar (horní), nebo Both. Výchozí Both.

.PARAMETER Mode
    Merge doplní chybějící položky (výchozí). Replace smaže cílovou navigaci
    a postaví ji znovu.

.PARAMETER Apply
    Provede změny. Bez něj skript jen vypíše, co by udělal.

.PARAMETER SkipTitles
    Položky, které se nekopírují. Výchozí je prázdný seznam, tedy kopíruje se
    všechno - duplicitám brání to, že položka se stejným názvem se nepřidá
    znovu.

.EXAMPLE
    # 1) Nejdřív se podívat, co by se stalo
    ./src/Copy-SiteNavigation.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/Project01" -TargetSiteUrl "https://contoso.sharepoint.com/sites/Project03"

.EXAMPLE
    # 2) Teprve pak přenést
    ./src/Copy-SiteNavigation.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/Project01" -TargetSiteUrl "https://contoso.sharepoint.com/sites/Project03" -Apply

.EXAMPLE
    # Postavit cílovou navigaci celou znovu (maže!)
    ./src/Copy-SiteNavigation.ps1 -SourceSiteUrl "..." -TargetSiteUrl "..." -Mode Replace -Apply

.NOTES
    Postup a řešení chyb: docs/08-kopirovani-navigace.md

    Vyžaduje PnP.PowerShell a práva vlastníka na obou webech.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $SourceSiteUrl,
    [Parameter(Mandatory = $true)] [string] $TargetSiteUrl,

    [string] $ClientId = $env:PNP_CLIENT_ID,
    [string] $OutputFolder = "./export",

    [ValidateSet("QuickLaunch", "TopNavigationBar", "Both")]
    [string] $Location = "Both",

    [ValidateSet("Merge", "Replace")]
    [string] $Mode = "Merge",

    [switch] $Apply,

    [int] $MaxDepth = 10,

    # Prázdné = kopíruje se všechno. Slučování už duplicitám brání tím, že
    # položku se stejným názvem nepřidá znovu. Odkaz jako "Documents" ze vzoru
    # míří na vlastní pohled knihovny, takže ho zahodit by byla chyba.
    [string[]] $SkipTitles = @()
)

$ErrorActionPreference = "Stop"

$script:Warnings = @()

function Write-Step($Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Add-NavWarning($Message) {
    $script:Warnings += $Message
    Write-Warning $Message
}

function Write-WarningSummary($OutFolder) {
    if ($script:Warnings.Count -gt 0) {
        $script:Warnings | Set-Content -Path "$OutFolder/navigation-warnings.txt" -Encoding UTF8
        Write-Host ""
        Write-Host "$($script:Warnings.Count) varování - podrobnosti v $OutFolder/navigation-warnings.txt" -ForegroundColor Yellow
    }
}

function Assert-Prerequisites {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "Modul PnP.PowerShell není nainstalovaný. Spusťte: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    if (-not $ClientId) {
        throw "Chybí ClientId. Zadejte -ClientId nebo nastavte PNP_CLIENT_ID. Podrobněji v docs/06-jak-spustit-export.md"
    }
    if ($SourceSiteUrl.TrimEnd("/") -eq $TargetSiteUrl.TrimEnd("/")) {
        throw "Zdrojový a cílový web jsou tentýž. Zkontrolujte -SourceSiteUrl a -TargetSiteUrl."
    }
}

function Get-LocationsToProcess($Requested) {
    if ($Requested -eq "Both") { return @("QuickLaunch", "TopNavigationBar") }
    return @($Requested)
}

# Server-relativní cesta webu, tedy část URL za doménou. Podle ní se poznává,
# které odkazy míří do zdrojového webu a je potřeba je přepsat.
function Get-ServerRelativePath($SiteUrl) {
    return ([uri]$SiteUrl).AbsolutePath.TrimEnd("/")
}

# Odkaz přepíše ze zdrojového webu na cílový. Odkazy mimo zdrojový web se
# nechávají být - typicky vedou na intranet nebo do jiné aplikace.
function Convert-NavigationUrl($Url, $SourcePath, $TargetPath) {
    if (-not $Url) { return $Url }

    # Pozor: [String]::Replace je obyčejná záměna podřetězce, ne regulární výraz.
    # Původní skript sem posílal [Regex]::Escape(), což je chyba - u cesty
    # obsahující tečku nebo pomlčku by záměna přestala fungovat.
    if ($Url.StartsWith($SourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $TargetPath + $Url.Substring($SourcePath.Length)
    }

    if ($Url -like "*$SourcePath*") {
        $index = $Url.IndexOf($SourcePath, [System.StringComparison]::OrdinalIgnoreCase)
        return $Url.Substring(0, $index) + $TargetPath + $Url.Substring($index + $SourcePath.Length)
    }

    return $Url
}

function Test-IsAbsoluteUrl($Url) {
    return $Url -match '^https?://'
}

# Načte navigaci včetně zanoření. Children se dotahují zvlášť, protože v
# seznamu z Get-PnPNavigationNode nejsou naplněné.
function Read-NavigationNode($Node, $Connection, $Depth, $MaxDepth) {
    $children = @()

    if ($Depth -lt $MaxDepth) {
        try {
            $detail = Get-PnPNavigationNode -Id $Node.Id -Connection $Connection -ErrorAction Stop

            $childNodes = $detail.Children
            if ($null -eq $childNodes) {
                $childNodes = Get-PnPProperty -ClientObject $detail -Property Children -Connection $Connection
            }

            foreach ($child in $childNodes) {
                $children += Read-NavigationNode $child $Connection ($Depth + 1) $MaxDepth
            }
        }
        catch {
            Add-NavWarning "Podřízené položky '$($Node.Title)' nelze přečíst: $($_.Exception.Message)"
        }
    }

    return [pscustomobject]@{
        Id         = $Node.Id
        Title      = $Node.Title
        Url        = $Node.Url
        IsExternal = $Node.IsExternal
        Depth      = $Depth
        Children   = @($children)
    }
}

function Read-Navigation($NavLocation, $Connection, $MaxDepth) {
    $roots = Get-PnPNavigationNode -Location $NavLocation -Connection $Connection

    $tree = foreach ($root in $roots) {
        Read-NavigationNode $root $Connection 1 $MaxDepth
    }
    return @($tree)
}

function Measure-NavigationNodes($Nodes) {
    $count = 0
    foreach ($node in $Nodes) {
        $count++
        $count += Measure-NavigationNodes $node.Children
    }
    return $count
}

function Show-NavigationTree($Nodes, $Indent = "  ") {
    foreach ($node in $Nodes) {
        $flag = if ($node.IsExternal) { "  (externí)" } else { "" }
        Write-Host "$Indent- $($node.Title)$flag"
        Write-Host "$Indent    $($node.Url)" -ForegroundColor DarkGray
        Show-NavigationTree $node.Children ($Indent + "    ")
    }
}

# Ploché řádky pro plán a CSV.
function Get-NavigationRows($Nodes, $SourcePath, $TargetPath, $ParentPath = "") {
    foreach ($node in $Nodes) {
        $path = if ($ParentPath) { "$ParentPath / $($node.Title)" } else { $node.Title }
        $newUrl = Convert-NavigationUrl $node.Url $SourcePath $TargetPath

        [pscustomobject]@{
            Path       = $path
            Depth      = $node.Depth
            Title      = $node.Title
            SourceUrl  = $node.Url
            TargetUrl  = $newUrl
            Rewritten  = $newUrl -ne $node.Url
            IsExternal = $node.IsExternal
        }

        Get-NavigationRows $node.Children $SourcePath $TargetPath $path
    }
}

function Remove-Navigation($NavLocation, $Connection) {
    $nodes = Get-PnPNavigationNode -Location $NavLocation -Connection $Connection

    foreach ($node in $nodes) {
        try {
            Remove-PnPNavigationNode -Identity $node.Id -Force -Connection $Connection
            Write-Host "  - $($node.Title)" -ForegroundColor Red
        }
        catch {
            Add-NavWarning "Položku '$($node.Title)' nelze odebrat: $($_.Exception.Message)"
        }
    }
}

# Přenese jednu úroveň a rekurzivně její potomky. V režimu Merge se položka se
# stejným názvem nepřidává znovu, ale zanoří se do ní a doplní se, co v ní
# chybí - proto se dá skript pouštět opakovaně, i když už část navigace existuje.
function Copy-NavigationLevel {
    param(
        $SourceNodes,
        $TargetNodes,
        $NavLocation,
        $TargetConnection,
        $ParentId,
        $SourcePath,
        $TargetPath,
        $Indent = "  "
    )

    $added = 0

    foreach ($node in $SourceNodes) {
        if ($SkipTitles -contains $node.Title) {
            Write-Host "$Indent. $($node.Title) (přeskočeno)" -ForegroundColor DarkGray
            continue
        }

        $match = @($TargetNodes | Where-Object { $_.Title -eq $node.Title })

        if ($match.Count -gt 0) {
            Write-Host "$Indent= $($node.Title)" -ForegroundColor DarkGray

            if ($node.Children.Count -gt 0) {
                $added += Copy-NavigationLevel -SourceNodes $node.Children `
                    -TargetNodes $match[0].Children `
                    -NavLocation $NavLocation `
                    -TargetConnection $TargetConnection `
                    -ParentId $match[0].Id `
                    -SourcePath $SourcePath `
                    -TargetPath $TargetPath `
                    -Indent ($Indent + "    ")
            }
            continue
        }

        $url = Convert-NavigationUrl $node.Url $SourcePath $TargetPath

        try {
            $arguments = @{
                Title      = $node.Title
                Url        = $url
                Location   = $NavLocation
                Connection = $TargetConnection
            }
            if ($ParentId) { $arguments["Parent"] = $ParentId }
            if ($node.IsExternal -or (Test-IsAbsoluteUrl $url)) { $arguments["External"] = $true }

            $created = Add-PnPNavigationNode @arguments -ErrorAction Stop
            $added++
            Write-Host "$Indent+ $($node.Title)" -ForegroundColor Green

            if ($node.Children.Count -gt 0) {
                if ($created -and $created.Id) {
                    $added += Copy-NavigationLevel -SourceNodes $node.Children `
                        -TargetNodes @() `
                        -NavLocation $NavLocation `
                        -TargetConnection $TargetConnection `
                        -ParentId $created.Id `
                        -SourcePath $SourcePath `
                        -TargetPath $TargetPath `
                        -Indent ($Indent + "    ")
                }
                else {
                    Add-NavWarning "U '$($node.Title)' se nepodařilo zjistit Id, podřízené položky se nepřenesly."
                }
            }
        }
        catch {
            Add-NavWarning "Položku '$($node.Title)' nelze přidat: $($_.Exception.Message)"
        }
    }

    return $added
}

Assert-Prerequisites

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

$sourcePath = Get-ServerRelativePath $SourceSiteUrl
$targetPath = Get-ServerRelativePath $TargetSiteUrl

Write-Step "Připojuji se ke zdrojovému webu"
Write-Host "  $SourceSiteUrl  ($sourcePath)"
$sourceConnection = Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

Write-Step "Připojuji se k cílovému webu"
Write-Host "  $TargetSiteUrl  ($targetPath)"
$targetConnection = Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

$locations = Get-LocationsToProcess $Location
$allRows = @()

foreach ($navLocation in $locations) {
    Write-Step "Zdrojová navigace: $navLocation"

    $sourceTree = Read-Navigation $navLocation $sourceConnection $MaxDepth
    $sourceCount = Measure-NavigationNodes $sourceTree

    if ($sourceCount -eq 0) {
        Write-Host "  prázdná, není co přenášet"
        continue
    }

    Show-NavigationTree $sourceTree
    Write-Host "  celkem položek: $sourceCount"

    $targetTree = Read-Navigation $navLocation $targetConnection $MaxDepth
    $targetTitles = @($targetTree | ForEach-Object { $_.Title })
    Write-Host "  v cíli už je: $($targetTitles.Count) položek na první úrovni: $($targetTitles -join ', ')"

    $rows = @(Get-NavigationRows $sourceTree $sourcePath $targetPath)
    foreach ($row in $rows) { $row | Add-Member -NotePropertyName Location -NotePropertyValue $navLocation }
    $allRows += $rows

    $rewritten = @($rows | Where-Object Rewritten).Count
    Write-Host "  odkazů k přepsání na cílový web: $rewritten z $($rows.Count)"
}

if ($allRows.Count -eq 0) {
    Write-Step "Zdrojová navigace je prázdná - konec"
    Write-WarningSummary $OutputFolder
    return
}

$allRows | Select-Object Location, Path, Depth, Title, SourceUrl, TargetUrl, Rewritten, IsExternal |
    Export-Csv -Path "$OutputFolder/navigation-plan.csv" -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  plán uložen do $OutputFolder/navigation-plan.csv"

if (-not $Apply) {
    Write-Step "DRY-RUN - v cíli se nic nezměnilo"
    Write-Host "Až bude plán v pořádku, spusťte stejný příkaz s přepínačem -Apply." -ForegroundColor Yellow
    if ($Mode -eq "Replace") {
        Write-Host "Pozor: -Mode Replace nejdřív SMAŽE celou cílovou navigaci." -ForegroundColor Red
    }
    Write-WarningSummary $OutputFolder
    return
}

foreach ($navLocation in $locations) {
    $sourceTree = Read-Navigation $navLocation $sourceConnection $MaxDepth
    if ((Measure-NavigationNodes $sourceTree) -eq 0) { continue }

    if ($Mode -eq "Replace") {
        Write-Step "Mažu cílovou navigaci: $navLocation"
        Remove-Navigation $navLocation $targetConnection
        $targetTree = @()
    }
    else {
        $targetTree = Read-Navigation $navLocation $targetConnection $MaxDepth
    }

    Write-Step "Přenáším navigaci: $navLocation"
    $added = Copy-NavigationLevel -SourceNodes $sourceTree `
        -TargetNodes $targetTree `
        -NavLocation $navLocation `
        -TargetConnection $targetConnection `
        -ParentId $null `
        -SourcePath $sourcePath `
        -TargetPath $targetPath

    Write-Host "  přidáno položek: $added"
}

Write-WarningSummary $OutputFolder
Write-Step "Hotovo"
