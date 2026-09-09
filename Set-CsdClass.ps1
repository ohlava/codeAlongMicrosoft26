param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ContentName,

    [Parameter(Mandatory = $false)]
    [string]$CsdClass
)

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " Set CSD Class" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Site:    $SiteUrl"
Write-Host "Content: $ContentName"

if ($CsdClass) {
    Write-Host "CSD:     $CsdClass"
}
else {
    Write-Host "CSD:     Será selecionado pelo utilizador"
}

Write-Host ""

# ---------------------------------------------------------
# Configuration
# ---------------------------------------------------------

$CsdFieldInternalName = "RevIMBCS"
$CsdFieldTitle = "CSD Class"

$ClientId = "CLIENT_ID"

# ---------------------------------------------------------
# 1. Connect to SharePoint
# ---------------------------------------------------------

Write-Host "A ligar ao SharePoint..." -ForegroundColor Cyan

try {

    Connect-PnPOnline `
        -Url $SiteUrl `
        -Interactive `
        -ClientId $ClientId `
        -ErrorAction Stop

}
catch {

    Write-Host ""
    Write-Host "ERRO: Não foi possível ligar ao SharePoint." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

Write-Host "Ligação estabelecida." -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------
# 2. Get Content
# ---------------------------------------------------------

Write-Host "A procurar o Content '$ContentName'..." -ForegroundColor Cyan

try {

    $list = Get-PnPList `
        -Identity $ContentName `
        -ErrorAction Stop

}
catch {

    Write-Host ""
    Write-Host "ERRO: O Content '$ContentName' não foi encontrado." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

Write-Host "Content encontrado: $($list.Title)" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------
# 3. Get fields
# ---------------------------------------------------------

Write-Host "A procurar as colunas do Content..." -ForegroundColor Cyan

try {

    $fields = Get-PnPProperty `
        -ClientObject $list `
        -Property Fields `
        -ErrorAction Stop

}
catch {

    Write-Host ""
    Write-Host "ERRO: Não foi possível obter as colunas do Content." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

# ---------------------------------------------------------
# 4. Find CSD Class field
# ---------------------------------------------------------

$taxonomyField = $fields |
    Where-Object {
        $_.InternalName -eq $CsdFieldInternalName
    } |
    Select-Object -First 1

if (-not $taxonomyField) {

    Write-Host ""
    Write-Host "O Content '$ContentName' não possui a coluna '$CsdFieldTitle'." `
        -ForegroundColor Yellow

    Write-Host "InternalName esperado: $CsdFieldInternalName" `
        -ForegroundColor Yellow

    exit 0
}

Write-Host "Campo encontrado:" -ForegroundColor Green
Write-Host "  Title:        $($taxonomyField.Title)"
Write-Host "  InternalName: $($taxonomyField.InternalName)"
Write-Host "  Type:         $($taxonomyField.TypeAsString)"
Write-Host "  TermSetId:    $($taxonomyField.TermSetId)"
Write-Host ""

# ---------------------------------------------------------
# 5. Get Term Store / Term Set
# ---------------------------------------------------------

Write-Host "A obter os valores disponíveis para '$CsdFieldTitle'..." `
    -ForegroundColor Cyan

try {

    $ctx = Get-PnPContext

    # Get Taxonomy Session
    $taxonomySession =
        [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)

    # Get default site collection term store
    $termStore =
        $taxonomySession.GetDefaultSiteCollectionTermStore()

    # Get the Term Set configured in the CSD field
    $termSet =
        $termStore.GetTermSet(
            [Guid]$taxonomyField.TermSetId
        )

    # Get all terms
    $terms = $termSet.GetAllTerms()

    $ctx.Load($terms)
    $ctx.ExecuteQuery()

}
catch {

    Write-Host ""
    Write-Host "ERRO: Não foi possível obter os valores do Term Set." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

if (-not $terms -or $terms.Count -eq 0) {

    Write-Host ""
    Write-Host "ERRO: O Term Set não contém valores." `
        -ForegroundColor Red

    exit 1
}

Write-Host "Valores encontrados: $($terms.Count)" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------
# 6. Select CSD Class
# ---------------------------------------------------------
#
# Se -CsdClass foi passado:
#     -> usa diretamente esse valor
#
# Se não foi passado:
#     -> mostra todos os termos numerados
#     -> utilizador escolhe uma opção
#
# ---------------------------------------------------------

if (-not $CsdClass) {

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " Valores disponíveis - CSD Class" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    $termOptions = @($terms)

    # -----------------------------------------------------
    # Show only the CSD value
    # -----------------------------------------------------

    foreach ($term in $termOptions) {

        Write-Host "  $($term.Name)"
    }

    Write-Host ""
    Write-Host "  0 Cancelar" -ForegroundColor Yellow
    Write-Host ""

    # -----------------------------------------------------
    # Ask user for the CSD value
    # -----------------------------------------------------

    $term = $null

    do {

        $selection = Read-Host "Escolha o CSD Class"

        # Cancel
        if ($selection -eq "0") {

            Write-Host ""
            Write-Host "Operação cancelada pelo utilizador." `
                -ForegroundColor Yellow

            exit 0
        }

        # Find exact CSD value
        $term = $termOptions |
            Where-Object {

                if ($_.Name -match '^(\d+(?:\.\d+)*)\b') {

                    $Matches[1] -eq $selection
                }
                else {

                    $_.Name -eq $selection
                }

            } |
            Select-Object -First 1

        if (-not $term) {

            Write-Host ""
            Write-Host "CSD '$selection' não encontrado." `
                -ForegroundColor Yellow

            Write-Host "Escolha um dos valores apresentados." `
                -ForegroundColor Yellow

            Write-Host ""
        }

    } while (-not $term)

    $CsdClass = $term.Name

}
else {

    # -----------------------------------------------------
    # CsdClass was provided directly
    # -----------------------------------------------------

    $term = $terms |
        Where-Object {
            $_.Name -eq $CsdClass
        } |
        Select-Object -First 1

    if (-not $term) {

        Write-Host ""
        Write-Host "ERRO: O termo '$CsdClass' não existe no Term Set." `
            -ForegroundColor Red

        Write-Host ""
        Write-Host "Term Set: $($taxonomyField.TermSetId)" `
            -ForegroundColor Yellow

        Write-Host ""
        Write-Host "Valores disponíveis:" -ForegroundColor Yellow

        foreach ($availableTerm in $terms) {

            Write-Host "  $($availableTerm.Name)"
        }

        exit 1
    }
}

# ---------------------------------------------------------
# 7. Selected term information
# ---------------------------------------------------------

$termGuid = $term.Id.ToString()

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host " CSD Class selecionado" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host "Name:     $($term.Name)"
Write-Host "TermGuid: $termGuid"
Write-Host ""

# ---------------------------------------------------------
# 8. Get all items
# ---------------------------------------------------------

Write-Host "A procurar os itens..." -ForegroundColor Cyan

try {

    $items = Get-PnPListItem `
        -List $list `
        -PageSize 100 `
        -Fields "FileLeafRef", $CsdFieldInternalName `
        -ErrorAction Stop

}
catch {

    Write-Host ""
    Write-Host "ERRO: Não foi possível obter os itens." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

Write-Host "Itens encontrados: $($items.Count)" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------
# 9. Confirmation
# ---------------------------------------------------------

Write-Host "=========================================" -ForegroundColor Yellow
Write-Host " Confirmação" -ForegroundColor Yellow
Write-Host "=========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Site:              $SiteUrl"
Write-Host "Content:           $ContentName"
Write-Host "CSD Class:         $CsdClass"
Write-Host "TermGuid:          $termGuid"
Write-Host "Itens a alterar:   $($items.Count)"
Write-Host ""

Write-Host "ATENÇÃO: Esta operação irá alterar a coluna" `
    -ForegroundColor Yellow

Write-Host "'$CsdFieldTitle' em TODOS os itens deste Content." `
    -ForegroundColor Yellow

Write-Host ""

$confirmation = Read-Host "Deseja continuar? (S/N)"

if ($confirmation -notmatch '^[sS]$') {

    Write-Host ""
    Write-Host "Operação cancelada pelo utilizador." `
        -ForegroundColor Yellow

    exit 0
}

Write-Host ""

# ---------------------------------------------------------
# 10. Update every item
# ---------------------------------------------------------

$updated = 0
$failed = 0

foreach ($item in $items) {

    $fileName = $item["FileLeafRef"]

    Write-Host "Atualizando item $($item.Id): $fileName"

    try {

        # Create taxonomy value
        $newValue = New-Object `
            Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue

        $newValue.Label = $CsdClass
        $newValue.TermGuid = $termGuid
        $newValue.WssId = -1

        # Set taxonomy value
        $taxonomyField.SetFieldValueByValue(
            $item,
            $newValue
        )

        # Mark item for update
        $item.Update()

        $updated++

    }
    catch {

        $failed++

        Write-Host "  ERRO no item $($item.Id): $($_.Exception.Message)" `
            -ForegroundColor Red
    }
}

# ---------------------------------------------------------
# 11. Save changes
# ---------------------------------------------------------

Write-Host ""
Write-Host "A gravar alterações no SharePoint..." -ForegroundColor Cyan

try {

    Invoke-PnPQuery

}
catch {

    Write-Host ""
    Write-Host "ERRO: Não foi possível gravar as alterações." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    exit 1
}

# ---------------------------------------------------------
# 12. Result
# ---------------------------------------------------------

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host " Concluído" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host "Site:              $SiteUrl"
Write-Host "Content:           $ContentName"
Write-Host "CSD Class:         $CsdClass"
Write-Host "TermGuid:          $termGuid"
Write-Host "Itens encontrados: $($items.Count)"
Write-Host "Itens atualizados: $updated"
Write-Host "Itens com erro:    $failed"
Write-Host ""