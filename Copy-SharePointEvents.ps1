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

$Tenant = "volkswagengroup.onmicrosoft.com"

$ShouldCopyValues = ($CopyValues -eq "yes")

# SharePoint Events / Calendar list template
$EventBaseTemplate = 106

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

# ============================================================
# SYSTEM FIELDS
#
# These are SharePoint internal fields which should not be
# recreated as custom fields.
#
# Calendar-specific fields are intentionally NOT included here.
# They need to remain available for Events.
# ============================================================

$SystemFields = @(
    "ID",
    "ContentType",
    "Modified",
    "Created",
    "Author",
    "Editor",
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
    "SelectTitle",
    "InstanceID",
    "Order",
    "FileSystemObjectType",
    "MetaInfo",
    "owshiddenversion",
    "WorkflowVersion",
    "WorkflowInstanceID",
    "EditMenuTableStart",
    "EditMenuTableEnd",
    "_EditMenuTableStart",
    "_EditMenuTableEnd",
    "LinkFilename",
    "LinkFilenameNoMenu",
    "LinkTitle",
    "LinkTitleNoMenu",
    "HTML_x0020_File_x0020_Type",
    "HTML_x0020_x0020_File_x0020_Type"
)

# ============================================================
# CALENDAR FIELDS
#
# These fields are important for SharePoint Events.
# They must be copied and remain visible.
# ============================================================

$CalendarFields = @(
    "EventDate",
    "EndDate",
    "fAllDayEvent",
    "fRecurrence",
    "RecurrenceData",
    "Recurrence",
    "Category",
    "Location",
    "Description",
    "UID",
    "EventType",
    "Workspace",
    "MasterSeriesItemID",
    "Duration",
    "TimeZone",
    "BannerUrl",
    "BannerImageUrl"
)

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
# HELPER - CHECK TARGET FIELD
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
# HELPER - CREATE TARGET FIELD
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

    $FieldName    = $SourceField.Title
    $SchemaXml    = $SourceField.SchemaXml

    if ([string]::IsNullOrWhiteSpace($SchemaXml)) {
        throw "SchemaXml is empty for field '$FieldName'."
    }

    # --------------------------------------------------------
    # Always reload target list before changing it.
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
# HELPER - UPDATE DEFAULT VIEW
#
# EVERYTHING copied from Site A is made visible.
#
# Additionally:
#   Modified
#   Modified By
#
# are always visible.
# ============================================================

function Set-TargetListView {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ListId,

        [Parameter(Mandatory = $true)]
        $Connection
    )

    try {

        Write-Host ""
        Write-Host "Updating default view..." -ForegroundColor Cyan

        # ----------------------------------------------------
        # Reload list
        # ----------------------------------------------------

        $FreshList = Get-PnPList `
            -Identity $ListId `
            -Connection $Connection `
            -Includes Id,Title `
            -ErrorAction Stop

        # ----------------------------------------------------
        # Read ALL target fields
        # ----------------------------------------------------

        $TargetFields = @(
            Get-PnPField `
                -List $FreshList.Id `
                -Connection $Connection `
                -Includes `
                    InternalName,
                    Title,
                    Hidden,
                    ReadOnlyField `
                -ErrorAction Stop
        )

        # ----------------------------------------------------
        # Make every non-system field visible.
        #
        # IMPORTANT:
        # For Events we intentionally include calendar fields.
        # ----------------------------------------------------

        $ViewFields = @()

        foreach ($Field in $TargetFields) {

            $InternalName = $Field.InternalName

            if ([string]::IsNullOrWhiteSpace($InternalName)) {
                continue
            }

            # Never put these internal SharePoint fields in view
            if ($SystemFields -contains $InternalName) {
                continue
            }

            # Add all fields that exist on the target
            if ($ViewFields -notcontains $InternalName) {

                $ViewFields += $InternalName
            }
        }

        # ----------------------------------------------------
        # Modified
        # ----------------------------------------------------

        if ($ViewFields -notcontains "Modified") {
            $ViewFields += "Modified"
        }

        # ----------------------------------------------------
        # Modified By
        # ----------------------------------------------------

        if ($ViewFields -notcontains "Editor") {
            $ViewFields += "Editor"
        }

        # ----------------------------------------------------
        # Get default view
        # ----------------------------------------------------

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

            $DefaultView = $Views |
                Where-Object {
                    $_.Title -eq "All Events"
                } |
                Select-Object -First 1
        }

        if ($null -eq $DefaultView) {
            throw "Could not find the default Events view."
        }

        Write-Host "  View: $($DefaultView.Title)" `
            -ForegroundColor Yellow

        # ----------------------------------------------------
        # Update view
        # ----------------------------------------------------

        Set-PnPView `
            -List $FreshList.Id `
            -Identity $DefaultView.Id `
            -Fields $ViewFields `
            -Connection $Connection `
            -ErrorAction Stop

        Write-Host "  Default view updated successfully." `
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
# HELPER - CONVERT EVENT FIELD VALUE
# ============================================================

function Convert-EventFieldValue {
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

        "Text" {
            return [string]$Value
        }

        "Note" {
            return [string]$Value
        }

        "MultiLineText" {
            return [string]$Value
        }

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

        "Choice" {
            return [string]$Value
        }

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

        "Number" {
            return $Value
        }

        "Currency" {
            return $Value
        }

        "Boolean" {
            return [bool]$Value
        }

        "DateTime" {

            if ($Value -is [DateTime]) {
                return $Value
            }

            try {
                return [DateTime]$Value
            }
            catch {
                return $Value
            }
        }

        "User" {

            if ($Value.Email) {
                return [string]$Value.Email
            }

            if ($Value.LookupValue) {
                return [string]$Value.LookupValue
            }

            return $null
        }

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

        "Lookup" {

            if ($Value.LookupId) {
                return [string]$Value.LookupId
            }

            return $null
        }

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

        "TaxonomyFieldType" {

            if ($Value.TermGuid) {
                return [string]$Value.TermGuid
            }

            if ($Value.Label) {
                return [string]$Value.Label
            }

            return $null
        }

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

        default {
            return $Value
        }
    }
}

# ============================================================
# HEADER
# ============================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "             SHAREPOINT EVENT / CALENDAR COPY" -ForegroundColor Cyan
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

    Write-Host "  YES - calendars, columns, events and values" `
        -ForegroundColor Green
}
else {

    Write-Host "  NO - ONLY calendar structure and columns" `
        -ForegroundColor Yellow
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
# GET SOURCE EVENT LISTS
# ============================================================

Write-Host ""
Write-Host "Reading SOURCE Event lists..." -ForegroundColor Cyan

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
            ItemCount,
            RootFolder `
        -ErrorAction Stop |
        Where-Object {
            $_.Hidden -eq $false -and
            $_.BaseTemplate -eq $EventBaseTemplate
        }
)

Write-Host "Event lists found: $($SourceLists.Count)" `
    -ForegroundColor Green

if ($SourceLists.Count -eq 0) {

    Write-Host ""
    Write-Host "No Event / Calendar lists were found." `
        -ForegroundColor Yellow
}

# ============================================================
# PROCESS EACH EVENT LIST
# ============================================================

foreach ($SourceList in $SourceLists) {

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host "CALENDAR: $($SourceList.Title)" `
        -ForegroundColor Cyan

    Write-Host "============================================================" `
        -ForegroundColor Cyan

    # ========================================================
    # GET / CREATE TARGET CALENDAR
    # ========================================================

    $TargetList = $null

    try {

        $TargetList = Get-PnPList `
            -Identity $SourceList.Title `
            -Connection $TargetConnection `
            -Includes Id,Title,RootFolder,BaseTemplate `
            -ErrorAction Stop

        Write-Host "Target Calendar already exists." `
            -ForegroundColor Yellow

        Write-Host "NOTHING will be deleted." `
            -ForegroundColor Yellow

        $ListsAlreadyExist++
    }
    catch {

        try {

            Write-Host "Creating target Calendar..." `
                -ForegroundColor Yellow

            # ------------------------------------------------
            # IMPORTANT:
            # Template 106 creates an Events / Calendar list.
            # ------------------------------------------------

            $TargetList = New-PnPList `
                -Title $SourceList.Title `
                -Template Events `
                -OnQuickLaunch:$false `
                -Connection $TargetConnection `
                -ErrorAction Stop

            $ListsCreated++

            Write-Host "Target Calendar created." `
                -ForegroundColor Green

            Start-Sleep -Milliseconds 1500

            $TargetList = Get-PnPList `
                -Identity $SourceList.Title `
                -Connection $TargetConnection `
                -Includes Id,Title,RootFolder,BaseTemplate `
                -ErrorAction Stop
        }
        catch {

            $ListsFailed++

            Write-Host "Could not create target Calendar '$($SourceList.Title)': $($_.Exception.Message)" `
                -ForegroundColor Red

            continue
        }
    }

    # ========================================================
    # READ SOURCE FIELDS
    # ========================================================

    Write-Host ""
    Write-Host "Reading ALL fields from source calendar..." `
        -ForegroundColor Cyan

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

        Write-Host "Could not read fields from calendar '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    # ========================================================
    # DETERMINE FIELDS TO COPY
    #
    # Copy:
    #   - every visible custom field
    #   - every calendar field
    #
    # Do NOT copy standard SharePoint system fields.
    # ========================================================

    $FieldsToCopy = @(
        $SourceFields |
        Where-Object {

            $IsCalendarField =
                $CalendarFields -contains $_.InternalName

            $IsCustomField =
                (-not $_.Hidden) -and
                (-not $_.ReadOnlyField) -and
                (-not ($SystemFields -contains $_.InternalName))

            $IsCalendarField -or $IsCustomField
        } |
        Sort-Object InternalName -Unique
    )

    Write-Host "Fields to copy: $($FieldsToCopy.Count)" `
        -ForegroundColor Green

    # ========================================================
    # CREATE FIELDS
    # ========================================================

    foreach ($Field in $FieldsToCopy) {

        $FieldName    = $Field.Title
        $InternalName = $Field.InternalName
        $FieldType    = $Field.TypeAsString

        if ([string]::IsNullOrWhiteSpace($InternalName)) {
            continue
        }

        # ----------------------------------------------------
        # Check target field
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

            Start-Sleep -Milliseconds 700
        }
        catch {

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

    # ========================================================
    # UPDATE DEFAULT VIEW
    #
    # This happens BEFORE item copying.
    # All copied fields become visible.
    # ========================================================

    Set-TargetListView `
        -ListId $TargetList.Id `
        -Connection $TargetConnection

    # ========================================================
    # COPY VALUES = NO
    #
    # DO NOT READ EVENTS.
    # ========================================================

    if (-not $ShouldCopyValues) {

        Write-Host ""
        Write-Host "CopyValues = NO -> events will NOT be copied." `
            -ForegroundColor Yellow

        Write-Host "Calendar structure and columns completed." `
            -ForegroundColor Green

        continue
    }

    # ========================================================
    # REFRESH TARGET CALENDAR
    # ========================================================

    try {

        $TargetList = Get-PnPList `
            -Identity $TargetList.Id `
            -Connection $TargetConnection `
            -Includes Id,Title,RootFolder,BaseTemplate `
            -ErrorAction Stop
    }
    catch {

        Write-Host "Could not refresh target Calendar '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    # ========================================================
    # GET FIELDS TO LOAD
    #
    # Load all fields that we intend to copy.
    # ========================================================

    $FieldsToLoad = @(
        $FieldsToCopy |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.InternalName)
        } |
        Select-Object -ExpandProperty InternalName
    )

    Write-Host ""
    Write-Host "Reading ALL events..." -ForegroundColor Cyan

    # ========================================================
    # READ SOURCE EVENTS
    # ========================================================

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

        Write-Host "Events found: $($SourceItems.Count)" `
            -ForegroundColor Green
    }
    catch {

        Write-Host "Could not read events from '$($SourceList.Title)': $($_.Exception.Message)" `
            -ForegroundColor Red

        continue
    }

    if ($SourceItems.Count -eq 0) {

        Write-Host "No events to copy." `
            -ForegroundColor DarkGray

        continue
    }

    # ========================================================
    # READ TARGET FIELDS
    # ========================================================

    Write-Host "Reading target calendar fields..." `
        -ForegroundColor Cyan

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
    # TARGET FIELD MAP
    # ========================================================

    $TargetFieldMap = @{}

    foreach ($TargetField in $TargetFields) {

        if (-not [string]::IsNullOrWhiteSpace($TargetField.InternalName)) {

            $TargetFieldMap[$TargetField.InternalName] = $TargetField
        }
    }

    # ========================================================
    # COPY EVENTS
    # ========================================================

    $ItemNumber = 0

    foreach ($SourceItem in $SourceItems) {

        $ItemNumber++

        Write-Host ""
        Write-Host "  Event $ItemNumber / $($SourceItems.Count)" `
            -ForegroundColor Cyan

        $Values = @{}

        # ----------------------------------------------------
        # BUILD EVENT VALUES
        # ----------------------------------------------------

        foreach ($SourceField in $FieldsToCopy) {

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

                $ConvertedValue = Convert-EventFieldValue `
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
        # CREATE EVENT
        # ----------------------------------------------------

        try {

            $TargetItem = Add-PnPListItem `
                -List $TargetList.Id `
                -Values $Values `
                -Connection $TargetConnection `
                -ErrorAction Stop

            $ItemsCreated++

            Write-Host "    Event created -> Target ID $($TargetItem.Id)" `
                -ForegroundColor Green
        }
        catch {

            $ItemsFailed++

            Write-Host "    Could not create event: $($_.Exception.Message)" `
                -ForegroundColor Red

            continue
        }
    }

    # ========================================================
    # UPDATE VIEW AGAIN
    #
    # This guarantees that fields created/available after the
    # event copy are visible.
    # ========================================================

    Set-TargetListView `
        -ListId $TargetList.Id `
        -Connection $TargetConnection

    Write-Host ""
    Write-Host "Calendar '$($SourceList.Title)' completed." `
        -ForegroundColor Green
}

# ============================================================
# FINAL SUMMARY
# ============================================================

Write-Host ""
Write-Host "============================================================" `
    -ForegroundColor Cyan

Write-Host "                  EVENT COPY SUMMARY" `
    -ForegroundColor Cyan

Write-Host "============================================================" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "COPY VALUES" -ForegroundColor Yellow

if ($ShouldCopyValues) {

    Write-Host "  Mode:                YES - calendar + fields + events"
}
else {

    Write-Host "  Mode:                NO - calendar structure only"
}

Write-Host ""

Write-Host "CALENDARS" -ForegroundColor Yellow

Write-Host "  Created:             $ListsCreated"
Write-Host "  Already existed:     $ListsAlreadyExist"
Write-Host "  Failed:              $ListsFailed"

Write-Host ""

Write-Host "FIELDS" -ForegroundColor Yellow

Write-Host "  Created:             $FieldsCreated"
Write-Host "  Already existed:     $FieldsAlreadyExist"
Write-Host "  Failed:              $FieldsFailed"

Write-Host ""

Write-Host "EVENTS" -ForegroundColor Yellow

if ($ShouldCopyValues) {

    Write-Host "  Found:               $ItemsFound"
}
else {

    Write-Host "  Found:               0 (not read)"
}

Write-Host "  Created:             $ItemsCreated"
Write-Host "  Failed:              $ItemsFailed"

Write-Host ""
Write-Host "============================================================" `
    -ForegroundColor Cyan

Write-Host "                       FINISHED" `
    -ForegroundColor Cyan

Write-Host "============================================================" `
    -ForegroundColor Cyan