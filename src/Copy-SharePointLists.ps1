param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$TargetSiteUrl,

    [Parameter(Mandatory = $true, Position = 2)]
    [ValidateSet("yes", "no")]
    [string]$CopyValues,

    [Parameter(Mandatory = $true, Position = 3)]
    [string]$ClientId
)

# ============================================================
# CONFIGURATION
# ============================================================

$ErrorActionPreference = "Stop"

$Tenant   = "volkswagengroup.onmicrosoft.com"

$ShouldCopyValues = ($CopyValues -eq "yes")

# ============================================================
# COUNTERS
# ============================================================

$ListsCreated       = 0
$ListsAlreadyExist  = 0
$ListsFailed        = 0

$FieldsCreated      = 0
$FieldsAlreadyExist = 0
$FieldsFailed       = 0

$ItemsFound         = 0
$ItemsCreated       = 0
$ItemsFailed        = 0

$AttachmentsCopied  = 0
$AttachmentsFailed  = 0

# ============================================================
# SYSTEM FIELDS
# These fields must not be copied as custom fields.
# ============================================================

$SystemFields = @(
    "ID",
    "ContentType",
    "Modified",
    "Created",
    "Author",
    "Editor",
    "_UIVersionString",
    "_UIVersionString",
    "Attachments",
    "Edit",
    "DocIcon",
    "FileLeafRef",
    "FileRef",
    "FileDirRef",
    "FSObjType",
    "SortBehavior",
    "PermMask",
    "UniqueId",
    "GUID",
    "_ModerationStatus",
    "_ModerationComments",
    "Created_x0020_Date",
    "Modified_x0020_Date",
    "AppAuthor",
    "AppEditor",
    "ComplianceAssetId",
    "_ComplianceFlags",
    "_ComplianceTag",
    "_ComplianceTagWrittenTime",
    "_ComplianceTagUserId",
    "_HasCopyDestinations",
    "_CopySource",
    "_UIVersion",
    "_Level",
    "_IsCurrentVersion",
    "_ModerationStatus",
    "SelectTitle",
    "InstanceID",
    "Order",
    "FSObjType",
    "FileSystemObjectType",
    "MetaInfo",
    "owshiddenversion",
    "WorkflowVersion",
    "WorkflowInstanceID",
    "Attachments",
    "EditMenuTableStart",
    "EditMenuTableEnd",
    "_EditMenuTableStart",
    "_EditMenuTableEnd",
    "LinkFilename",
    "LinkFilenameNoMenu",
    "LinkTitle",
    "LinkTitleNoMenu",
    "HTML_x0020_File_x0020_Type",
    "HTML_x0020_x0020_File_x0020_Type",
    "_HasCopyDestinations",
    "_CopySource",
    "InstanceID",
    "FSObjType",
    "FileLeafRef",
    "FileRef"
)

function Set-TargetListView {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ListId,

        [Parameter(Mandatory = $true)]
        $Connection
    )

    try {
        # Reload the list
        $FreshList = Get-PnPList `
            -Identity $ListId `
            -Connection $Connection `
            -Includes Id,Title `
            -ErrorAction Stop

        # Get all fields from target list
        $TargetFields = @(
            Get-PnPField `
                -List $FreshList.Id `
                -Connection $Connection `
                -Includes InternalName,Hidden,ReadOnlyField `
                -ErrorAction Stop
        )

        # All fields that are not hidden/read-only/system fields
        # All copied fields that should be visible
        $ViewFields = @(
            $TargetFields |
            Where-Object {
                -not $_.Hidden -and
                -not $_.ReadOnlyField -and
                -not ($SystemFields -contains $_.InternalName)
            } |
            Select-Object -ExpandProperty InternalName
        )

        # Always show Modified and Modified By
        $MandatoryFields = @(
            "Modified",
            "Editor"
        )

        foreach ($MandatoryField in $MandatoryFields) {

            if ($ViewFields -notcontains $MandatoryField) {
                $ViewFields += $MandatoryField
            }
        }

        if ($ViewFields.Count -eq 0) {
            Write-Host "  No fields available for the view." `
                -ForegroundColor Yellow
            return
        }

        # Get default view
        $Views = @(
            Get-PnPView `
                -List $FreshList.Id `
                -Connection $Connection `
                -ErrorAction Stop
        )

        $DefaultView = $Views |
            Where-Object {
                $_.DefaultView -eq $true
            } |
            Select-Object -First 1

        if ($null -eq $DefaultView) {
            throw "Could not find the default view."
        }

        Write-Host "  Updating default view: $($DefaultView.Title)" `
            -ForegroundColor Yellow

        Set-PnPView `
            -List $FreshList.Id `
            -Identity $DefaultView.Id `
            -Fields $ViewFields `
            -Connection $Connection `
            -ErrorAction Stop

        Write-Host "  Default view updated." `
            -ForegroundColor Green

        Write-Host "  Visible fields: $($ViewFields.Count)" `
            -ForegroundColor Green
    }
    catch {
        Write-Host "  Could not update default view: $($_.Exception.Message)" `
            -ForegroundColor Red
    }
}

# ============================================================
# HELPER - GET FIELD TYPE
# ============================================================

function Get-FieldType {
    param(
        [Parameter(Mandatory = $true)]
        $Field
    )

    if ($null -ne $Field.TypeAsString) {
        return $Field.TypeAsString
    }

    return ""
}

# ============================================================
# HELPER - CHECK WHETHER FIELD EXISTS
# ============================================================

function Get-TargetField {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$InternalName,

        [Parameter(Mandatory = $true)]
        $Connection
    )

    try {
        $Fields = @(
            Get-PnPField `
                -List $ListId `
                -Connection $Connection `
                -Includes Id,Title,InternalName,StaticName,TypeAsString,Hidden,ReadOnlyField `
                -ErrorAction Stop
        )

        return $Fields |
            Where-Object {
                $_.InternalName -eq $InternalName
            } |
            Select-Object -First 1
    }
    catch {
        return $null
    }
}

# ============================================================
# HELPER - CREATE FIELD
# ============================================================

function Add-TargetField {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ListId,

        [Parameter(Mandatory = $true)]
        $SourceField,

        [Parameter(Mandatory = $true)]
        $Connection
    )

    $FieldName         = $SourceField.Title
    $InternalName      = $SourceField.InternalName
    $SchemaXml         = $SourceField.SchemaXml
    $TypeAsString      = $SourceField.TypeAsString

    if ([string]::IsNullOrWhiteSpace($SchemaXml)) {
        throw "SchemaXml is empty for field '$FieldName'."
    }

    # --------------------------------------------------------
    # IMPORTANT:
    #
    # Reload the target list immediately before modification.
    # SharePoint can reject a stale List object with:
    #
    # "The object has been updated by another user since it
    # was last fetched."
    # --------------------------------------------------------

    $FreshTargetList = Get-PnPList `
        -Identity $ListId `
        -Connection $Connection `
        -Includes Id,Title `
        -ErrorAction Stop

    Add-PnPFieldFromXml `
        -List $FreshTargetList.Id `
        -FieldXml $SchemaXml `
        -Connection $Connection `
        -ErrorAction Stop
}

# ============================================================
# HELPER - CONVERT ITEM FIELD VALUE
# ============================================================

function Convert-FieldValue {
    param(
        [Parameter(Mandatory = $true)]
        $Field,

        [Parameter(Mandatory = $false)]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $Type = $Field.TypeAsString

    switch ($Type) {

        # ----------------------------------------------------
        # TEXT / NOTE
        # ----------------------------------------------------

        "Text" {
            return [string]$Value
        }

        "Note" {
            return [string]$Value
        }

        "MultiLineText" {
            return [string]$Value
        }

        # ----------------------------------------------------
        # URL
        # ----------------------------------------------------

        "URL" {
            if ($Value.Url) {
                $Url = [string]$Value.Url

                $Description = ""

                if ($Value.Description) {
                    $Description = [string]$Value.Description
                }

                if ([string]::IsNullOrWhiteSpace($Description)) {
                    return $Url
                }

                return "$Url, $Description"
            }

            return [string]$Value
        }

        # ----------------------------------------------------
        # CHOICE
        # ----------------------------------------------------

        "Choice" {
            return [string]$Value
        }

        # ----------------------------------------------------
        # MULTI CHOICE
        # ----------------------------------------------------

        "MultiChoice" {

            if ($Value -is [System.Array]) {
                return @(
                    $Value |
                    ForEach-Object {
                        [string]$_
                    }
                )
            }

            return @([string]$Value)
        }

        # ----------------------------------------------------
        # NUMBER
        # ----------------------------------------------------

        "Number" {
            return $Value
        }

        # ----------------------------------------------------
        # CURRENCY
        # ----------------------------------------------------

        "Currency" {
            return $Value
        }

        # ----------------------------------------------------
        # BOOLEAN
        # ----------------------------------------------------

        "Boolean" {
            return [bool]$Value
        }

        # ----------------------------------------------------
        # DATE / DATETIME
        # ----------------------------------------------------

        "DateTime" {
            return $Value
        }

        # ----------------------------------------------------
        # USER
        # ----------------------------------------------------

        "User" {

            if ($Value.Email) {
                return [string]$Value.Email
            }

            if ($Value.LookupValue) {
                return [string]$Value.LookupValue
            }

            return $null
        }

        # ----------------------------------------------------
        # MULTI USER
        # ----------------------------------------------------

        "UserMulti" {

            $Result = @()

            if ($Value -is [System.Array]) {

                foreach ($User in $Value) {

                    if ($User.Email) {
                        $Result += [string]$User.Email
                    }
                    elseif ($User.LookupValue) {
                        $Result += [string]$User.LookupValue
                    }
                }
            }
            else {

                if ($Value.Email) {
                    $Result += [string]$Value.Email
                }
                elseif ($Value.LookupValue) {
                    $Result += [string]$Value.LookupValue
                }
            }

            return $Result
        }

        # ----------------------------------------------------
        # LOOKUP
        # ----------------------------------------------------

        "Lookup" {

            if ($Value.LookupId) {
                return [string]$Value.LookupId
            }

            return $null
        }

        # ----------------------------------------------------
        # MULTI LOOKUP
        # ----------------------------------------------------

        "LookupMulti" {

            $Result = @()

            if ($Value -is [System.Array]) {

                foreach ($Lookup in $Value) {

                    if ($Lookup.LookupId) {
                        $Result += [string]$Lookup.LookupId
                    }
                }
            }
            else {

                if ($Value.LookupId) {
                    $Result += [string]$Value.LookupId
                }
            }

            return $Result
        }

        # ----------------------------------------------------
        # TAXONOMY
        # ----------------------------------------------------

        "TaxonomyFieldType" {

            if ($Value.TermGuid) {
                return [string]$Value.TermGuid
            }

            if ($Value.Label) {
                return [string]$Value.Label
            }

            return $null
        }

        # ----------------------------------------------------
        # MULTI TAXONOMY
        # ----------------------------------------------------

        "TaxonomyFieldTypeMulti" {

            $Result = @()

            if ($Value -is [System.Array]) {

                foreach ($Term in $Value) {

                    if ($Term.TermGuid) {
                        $Result += [string]$Term.TermGuid
                    }
                    elseif ($Term.Label) {
                        $Result += [string]$Term.Label
                    }
                }
            }
            else {

                if ($Value.TermGuid) {
                    $Result += [string]$Value.TermGuid
                }
                elseif ($Value.Label) {
                    $Result += [string]$Value.Label
                }
            }

            return $Result
        }

        # ----------------------------------------------------
        # DEFAULT
        # ----------------------------------------------------

        default {
            return $Value
        }
    }
}

# ============================================================
# HELPER - COPY ATTACHMENTS
# ============================================================

function Copy-ItemAttachments {
    param(
        [Parameter(Mandatory = $true)]
        $SourceItem,

        [Parameter(Mandatory = $true)]
        $TargetItem,

        [Parameter(Mandatory = $true)]
        $SourceList,

        [Parameter(Mandatory = $true)]
        $TargetList,

        [Parameter(Mandatory = $true)]
        $SourceConnection,

        [Parameter(Mandatory = $true)]
        $TargetConnection
    )

    $TempDirectory = Join-Path `
        ([System.IO.Path]::GetTempPath()) `
        ("SharePointCopy_" + [Guid]::NewGuid().ToString())

    try {

        New-Item `
            -ItemType Directory `
            -Path $TempDirectory `
            -Force `
            | Out-Null

        $AttachmentFolder = "$($SourceList.RootFolder.ServerRelativeUrl)/Attachments/$($SourceItem.Id)"

        $Attachments = @(
            Get-PnPFolderItem `
                -FolderSiteRelativeUrl $AttachmentFolder `
                -Connection $SourceConnection `
                -ItemType File `
                -ErrorAction SilentlyContinue
        )

        foreach ($Attachment in $Attachments) {

            try {

                $LocalFile = Join-Path `
                    $TempDirectory `
                    $Attachment.Name

                Get-PnPFile `
                    -Url $Attachment.ServerRelativeUrl `
                    -Path $TempDirectory `
                    -FileName $Attachment.Name `
                    -AsFile `
                    -Force `
                    -Connection $SourceConnection `
                    -ErrorAction Stop

                Add-PnPListItemAttachment `
                    -List $TargetList.Id `
                    -ListItemId $TargetItem.Id `
                    -FilePath $LocalFile `
                    -Connection $TargetConnection `
                    -ErrorAction Stop

                $script:AttachmentsCopied++

                Write-Host "      Attachment copied: $($Attachment.Name)" `
                    -ForegroundColor DarkGreen
            }
            catch {

                $script:AttachmentsFailed++

                Write-Host "      Could not copy attachment '$($Attachment.Name)': $($_.Exception.Message)" `
                    -ForegroundColor Red
            }
        }
    }
    finally {

        if (Test-Path $TempDirectory) {

            Remove-Item `
                -Path $TempDirectory `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================
# HEADER
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "              SHAREPOINT LIST COPY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "SOURCE:" -ForegroundColor Yellow
Write-Host "  $SourceSiteUrl"

Write-Host ""
Write-Host "TARGET:" -ForegroundColor Yellow
Write-Host "  $TargetSiteUrl"

Write-Host ""
Write-Host "COPY VALUES:" -ForegroundColor Yellow

if ($ShouldCopyValues) {
    Write-Host "  YES - lists, columns, items and attachments" -ForegroundColor Green
}
else {
    Write-Host "  NO - ONLY list structure and columns will be copied" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# CONNECT SOURCE
# ============================================================

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan

$SourceConnection = Connect-PnPOnline `
    -Url $SourceSiteUrl `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -Interactive `
    -ReturnConnection

Write-Host "Connected to SOURCE." -ForegroundColor Green

# ============================================================
# CONNECT TARGET
# ============================================================

Write-Host ""
Write-Host "Connecting to TARGET..." -ForegroundColor Cyan

$TargetConnection = Connect-PnPOnline `
    -Url $TargetSiteUrl `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -Interactive `
    -ReturnConnection

Write-Host "Connected to TARGET." -ForegroundColor Green

# ============================================================
# GET SOURCE LISTS
# ============================================================

Write-Host ""
Write-Host "Reading SOURCE lists..." -ForegroundColor Cyan

$SourceLists = @(
    Get-PnPList `
        -Connection $SourceConnection `
        -Includes `
            Title,
            Id,
            BaseTemplate,
            BaseType,
            Hidden,
            Description,
            EnableAttachments,
            ItemCount `
        -ErrorAction Stop |
        Where-Object {
            $_.Hidden -eq $false -and
            $_.BaseTemplate -eq 100
        }
)

Write-Host "Source Lists found: $($SourceLists.Count)" -ForegroundColor Green

# ============================================================
# PROCESS EACH LIST
# ============================================================

foreach ($SourceList in $SourceLists) {

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "LIST: $($SourceList.Title)" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan

    # ========================================================
    # GET / CREATE TARGET LIST
    # ========================================================

    $TargetList = $null

    try {

        $TargetList = Get-PnPList `
            -Identity $SourceList.Title `
            -Connection $TargetConnection `
            -Includes Id,Title,RootFolder `
            -ErrorAction Stop

        Write-Host "Target List already exists. NOTHING will be deleted." `
            -ForegroundColor Yellow

        $ListsAlreadyExist++
    }
    catch {

        try {

            Write-Host "Creating target List..." -ForegroundColor Yellow
    
            $TargetList = New-PnPList `
            -Title $SourceList.Title `
            -Template GenericList `
            -OnQuickLaunch:$false `
            -Connection $TargetConnection `
            -ErrorAction Stop

            $ListsCreated++

            Write-Host "Target List created." -ForegroundColor Green

            Start-Sleep -Milliseconds 1000

            $TargetList = Get-PnPList `
                -Identity $SourceList.Title `
                -Connection $TargetConnection `
                -Includes Id,Title,RootFolder `
                -ErrorAction Stop
        }
        catch {

            $ListsFailed++

            Write-Host "Could not create target List '$($SourceList.Title)': $($_.Exception.Message)" `
                -ForegroundColor Red

            continue
        }
    }

    # ========================================================
    # READ SOURCE FIELDS
    # ========================================================

    Write-Host "Reading ALL fields from source..." -ForegroundColor Cyan

    try {

        $SourceFields = @(
            Get-PnPField `
                -List $SourceList.Id `
                -Connection $SourceConnection `
                -Includes `
                    Id,
                    Title,
                    InternalName,
                    StaticName,
                    TypeAsString,
                    Hidden,
                    ReadOnlyField,
                    Required,
                    Description,
                    Group,
                    SchemaXml,
                    DefaultValue `
                -ErrorAction Stop
        )

        Write-Host "Source fields found: $($SourceFields.Count)" `
            -ForegroundColor Green
    }
    catch {

        Write-Host "Could not read fields from '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    # ========================================================
    # FILTER VISIBLE CUSTOM FIELDS
    # ========================================================

    $VisibleSourceFields = @(
        $SourceFields |
        Where-Object {

            -not $_.Hidden -and
            -not $_.ReadOnlyField -and
            -not ($SystemFields -contains $_.InternalName)
        }
    )

    Write-Host "Visible custom fields: $($VisibleSourceFields.Count)" `
        -ForegroundColor Green

    # ========================================================
    # CREATE FIELDS
    # ========================================================

    foreach ($Field in $VisibleSourceFields) {

        $FieldName    = $Field.Title
        $InternalName = $Field.InternalName
        $FieldType    = $Field.TypeAsString

        # ----------------------------------------------------
        # Check if field already exists.
        #
        # IMPORTANT:
        # Fetch current target fields on every iteration.
        # ----------------------------------------------------

        $ExistingTargetField = Get-TargetField `
            -ListId $TargetList.Id `
            -InternalName $InternalName `
            -Connection $TargetConnection

        if ($ExistingTargetField) {

            Write-Host "  Field exists: $FieldName [$FieldType]" `
                -ForegroundColor DarkGray

            $FieldsAlreadyExist++

            continue
        }

        Write-Host "  Creating: $FieldName [$FieldType]" `
            -ForegroundColor Yellow

        try {

            Add-TargetField `
                -ListId $TargetList.Id `
                -SourceField $Field `
                -Connection $TargetConnection

            Write-Host "    Created successfully." `
                -ForegroundColor Green

            $FieldsCreated++

            # ------------------------------------------------
            # Give SharePoint a moment to commit the update.
            # ------------------------------------------------

            Start-Sleep -Milliseconds 700
        }
        catch {

            # ------------------------------------------------
            # Sometimes SharePoint returns the concurrency
            # error even after reloading the List.
            #
            # Retry once after refreshing everything.
            # ------------------------------------------------

            $ErrorMessage = $_.Exception.Message

            if ($ErrorMessage -match "updated by another user") {

                Write-Host "    SharePoint concurrency conflict. Retrying..." `
                    -ForegroundColor Yellow

                try {

                    Start-Sleep -Milliseconds 1500

                    $FreshTargetList = Get-PnPList `
                        -Identity $TargetList.Id `
                        -Connection $TargetConnection `
                        -Includes Id,Title `
                        -ErrorAction Stop

                    # Check again because another operation may
                    # have created the field already.

                    $ExistingAfterRetry = Get-TargetField `
                        -ListId $FreshTargetList.Id `
                        -InternalName $InternalName `
                        -Connection $TargetConnection

                    if ($ExistingAfterRetry) {

                        Write-Host "    Field now exists after refresh." `
                            -ForegroundColor Green

                        $FieldsAlreadyExist++

                        continue
                    }

                    Add-PnPFieldFromXml `
                        -List $FreshTargetList.Id `
                        -FieldXml $Field.SchemaXml `
                        -Connection $TargetConnection `
                        -ErrorAction Stop

                    Write-Host "    Created successfully on retry." `
                        -ForegroundColor Green

                    $FieldsCreated++

                    Start-Sleep -Milliseconds 1000
                }
                catch {

                    Write-Host "    Could not create field '$FieldName': $($_.Exception.Message)" `
                        -ForegroundColor Red

                    $FieldsFailed++
                }
            }
            else {

                Write-Host "    Could not create field '$FieldName': $ErrorMessage" `
                    -ForegroundColor Red

                $FieldsFailed++
            }
        }
    }

    Write-Host ""
    Write-Host "Updating default view..." -ForegroundColor Cyan

    Set-TargetListView `
    -ListId $TargetList.Id `
    -Connection $TargetConnection

    # ========================================================
    # COPY VALUES = NO
    #
    # IMPORTANT:
    # Do NOT read source items.
    # Do NOT create target items.
    # Do NOT process attachments.
    # ========================================================

    if (-not $ShouldCopyValues) {

        Write-Host ""
        Write-Host "CopyValues = NO -> items will NOT be copied." `
            -ForegroundColor Yellow

        Write-Host "List structure and columns completed." `
            -ForegroundColor Green

        continue
    }

    # ========================================================
    # REFRESH TARGET LIST BEFORE ITEMS
    # ========================================================

    try {

        $TargetList = Get-PnPList `
            -Identity $TargetList.Id `
            -Connection $TargetConnection `
            -Includes Id,Title,RootFolder `
            -ErrorAction Stop
    }
    catch {

        Write-Host "Could not refresh target List '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    # ========================================================
    # GET SOURCE FIELDS TO LOAD
    # ========================================================

    $FieldsToLoad = @(
        $VisibleSourceFields |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.InternalName)
        } |
        Select-Object -ExpandProperty InternalName
    )

    # ========================================================
    # READ SOURCE ITEMS
    # ========================================================

    Write-Host ""
    Write-Host "Reading ALL items..." -ForegroundColor Cyan

    try {

        $SourceItems = @(
            Get-PnPListItem `
                -List $SourceList.Id `
                -PageSize 5000 `
                -Fields $FieldsToLoad `
                -Connection $SourceConnection `
                -ErrorAction Stop
        )

        $ItemsFound += $SourceItems.Count

        Write-Host "Items found: $($SourceItems.Count)" `
            -ForegroundColor Green
    }
    catch {

        Write-Host "Could not read items from '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    if ($SourceItems.Count -eq 0) {

        Write-Host "No items to copy." -ForegroundColor DarkGray

        continue
    }

    # ========================================================
    # GET TARGET FIELDS
    # ========================================================

    Write-Host "Reading target fields..." -ForegroundColor Cyan

    try {

        $TargetFields = @(
            Get-PnPField `
                -List $TargetList.Id `
                -Connection $TargetConnection `
                -Includes `
                    Id,
                    Title,
                    InternalName,
                    StaticName,
                    TypeAsString,
                    Hidden,
                    ReadOnlyField `
                -ErrorAction Stop
        )
    }
    catch {

        Write-Host "Could not read target fields: $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    # ========================================================
    # BUILD TARGET FIELD MAP
    # ========================================================

    $TargetFieldMap = @{}

    foreach ($TargetField in $TargetFields) {

        if (-not [string]::IsNullOrWhiteSpace($TargetField.InternalName)) {

            $TargetFieldMap[$TargetField.InternalName] = $TargetField
        }
    }

    # ========================================================
    # COPY ITEMS
    # ========================================================

    $ItemNumber = 0

    foreach ($SourceItem in $SourceItems) {

        $ItemNumber++

        Write-Host ""
        Write-Host "  Item $ItemNumber / $($SourceItems.Count)" `
            -ForegroundColor Cyan

        $Values = @{}

        # ----------------------------------------------------
        # BUILD ITEM VALUES
        # ----------------------------------------------------

        foreach ($SourceField in $VisibleSourceFields) {

            $InternalName = $SourceField.InternalName

            if ([string]::IsNullOrWhiteSpace($InternalName)) {
                continue
            }

            if (-not $TargetFieldMap.ContainsKey($InternalName)) {
                continue
            }

            if (-not $SourceItem.FieldValues.ContainsKey($InternalName)) {
                continue
            }

            $Value = $SourceItem.FieldValues[$InternalName]

            if ($null -eq $Value) {
                continue
            }

            try {

                $ConvertedValue = Convert-FieldValue `
                    -Field $SourceField `
                    -Value $Value

                if ($null -ne $ConvertedValue) {

                    $Values[$InternalName] = $ConvertedValue
                }
            }
            catch {

                Write-Host "    Could not convert field '$($SourceField.Title)': $($_.Exception.Message)" `
                    -ForegroundColor Yellow
            }
        }

        # ----------------------------------------------------
        # CREATE TARGET ITEM
        # ----------------------------------------------------

        try {

            $TargetItem = Add-PnPListItem `
                -List $TargetList.Id `
                -Values $Values `
                -Connection $TargetConnection `
                -ErrorAction Stop

            $ItemsCreated++

            Write-Host "    Item created -> Target ID $($TargetItem.Id)" `
                -ForegroundColor Green
        }
        catch {

            $ItemsFailed++

            Write-Host "    Could not create item: $($_.Exception.Message)" `
                -ForegroundColor Red

            continue
        }

        # ----------------------------------------------------
        # COPY ATTACHMENTS
        # ----------------------------------------------------

        if ($SourceList.EnableAttachments) {

            try {

                Copy-ItemAttachments `
                    -SourceItem $SourceItem `
                    -TargetItem $TargetItem `
                    -SourceList $SourceList `
                    -TargetList $TargetList `
                    -SourceConnection $SourceConnection `
                    -TargetConnection $TargetConnection
            }
            catch {

                Write-Host "    Could not process attachments: $($_.Exception.Message)" `
                    -ForegroundColor Red

                $AttachmentsFailed++
            }
        }
    }

    Write-Host ""
    Write-Host "List '$($SourceList.Title)' completed." `
        -ForegroundColor Green
}

# ============================================================
# FINAL SUMMARY
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "                     COPY SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "COPY VALUES" -ForegroundColor Yellow

if ($ShouldCopyValues) {
    Write-Host "  Mode:                YES - structure + values + attachments"
}
else {
    Write-Host "  Mode:                NO - structure only"
}

Write-Host ""
Write-Host "LISTS" -ForegroundColor Yellow
Write-Host "  Created:             $ListsCreated"
Write-Host "  Already existed:     $ListsAlreadyExist"
Write-Host "  Failed:              $ListsFailed"

Write-Host ""
Write-Host "FIELDS" -ForegroundColor Yellow
Write-Host "  Created:             $FieldsCreated"
Write-Host "  Already existed:     $FieldsAlreadyExist"
Write-Host "  Failed:              $FieldsFailed"

Write-Host ""
Write-Host "ITEMS" -ForegroundColor Yellow

if ($ShouldCopyValues) {
    Write-Host "  Found:               $ItemsFound"
}
else {
    Write-Host "  Found:               0 (not read)"
}

Write-Host "  Created:             $ItemsCreated"
Write-Host "  Failed:              $ItemsFailed"

Write-Host ""
Write-Host "ATTACHMENTS" -ForegroundColor Yellow

if ($ShouldCopyValues) {
    Write-Host "  Copied:              $AttachmentsCopied"
}
else {
    Write-Host "  Copied:              0 (not processed)"
}

Write-Host "  Failed:              $AttachmentsFailed"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "                       FINISHED" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
