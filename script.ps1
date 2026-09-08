#Sctipt for copy from template#
Bitte beim Änderungen am Code diese kurz dokumentieren

#v2.6
add Banner-copy function
Config

$SiteDomain = "https://volkswagengroup.sharepoint.com"
$SourcePath = "/sites/Project01"
$TargetPath = "/sites/Project03"

$SetOfflineAvailable = "nein" #ja

##Copy SC > Subweb, Subweb > Subweb, Subweb > SC
#Set '$true' if you want to copy from SiteCollection to SubWeb. Set '$false' if you want to copy from SubWeb to SiteCollection
#Default = $false
$CopyFromMainToSubSite = $false

##Copy from list. If u wanna copy Vorlage in more sites
#Default = $false
$CopyFromList = $false
$XmlPathList = "$PSScriptRoot\Sites.xml"

####################

$LocalFolder = $PSScriptRoot #if there is a 'Path Error' please change it with a absolute path. like: "C:\Users\FBYCUN8\Downloads"
$CopyCount = 1000 #copy n-1 Items (for testing in a certain number of the list)
$IsDevMode = $false

####################

$SourceSiteUrl = "$SiteDomain$SourcePath"
$TargetSiteUrl = "$SiteDomain$TargetPath"
$FolderPath = "$LocalFolder\FileTemp"
$ReplaceTarget = $true;
$IsCopyPages = $true;
$IsCopyTemplateDesign = $true;
$IsCopyRegionalSettings = $true;
$IsCopyNavigation = $true;

try {

if ($SourceSiteUrl -and $TargetSiteUrl) {
    
    Write-Host "Connecting to Site :'$($SourcePath)'........" -ForegroundColor Yellow  
	
	$ClientId = "--------------will be modified-----------"
    $Con1 = Connect-PnPOnline -Url $SourceSiteUrl  -Interactive -ClientId $ClientId -ReturnConnection -WarningAction Ignore
    
    Write-Host "Connection Successfull to site: '$($SourcePath)'" -ForegroundColor Green             
    
}
else {
    Write-Host "Source Site URL is empty" -ForegroundColor Red
    Break
}

}
catch {
Write-Host "Error in connecting to Site:'$($SiteUrl)'" $_.Exception.Message -ForegroundColor Red
exit
}

$global:Dataa = @()

$global:CurrentSitePath

Function Copy-SPOAttachments($SourceItem, $TargetItem, $ListName) {
Try {

    $Attachments = Get-PnPProperty -ClientObject $SourceItem -Property "AttachmentFiles" -Connection $Con1
    $attachments | ForEach-Object { 
        $Attachment = Get-PnPFile -Url $_.ServerRelativeUrl -FileName $_.FileName -Path $FolderPath -AsFile -Force -Connection $Con1
        $AddAttachment = Add-PnPListItemAttachment -List $ListName -Identity $TargetItem.Id -Path ("$FolderPath\"+$_.FileName) -Connection $global:Con2
    } 
    
}
Catch {
    write-host -f Red "Error Copying Attachments:" $_.Exception.Message
}

}

Function Copy-SPOListItems() {
param
(
[Parameter(Mandatory = $true)] [string] $SourceListName,
[Parameter(Mandatory = $true)] [string] $TargetListName
)
Try {
#Get All Items from the Source List in batches
Write-Progress -Activity "Reading Source..." -Status "Getting Items from Source List. Please wait..."
$SourceListItems = Get-PnPListItem -List $SourceListName -PageSize 5000 -Connection $Con1
$SourceListItemsCount = $SourceListItems.count
Write-host "--------Copy "$SourceListItemsCount" items from Source list"

    #Get fields to Update from the Source List - Skip Read only, hidden fields, content type and attachments
    #$SourceListFields = Get-PnPField -Connection $Con1 -List $SourceListName | Where { (-Not ($_.Hidden)) -and ($_.InternalName -ne "ContentType") -and ($_.InternalName -ne "Attachments") }
    $SourceListFields = Get-PnPField -Connection $Con1 -List $SourceListName | Where { (-Not ($_.ReadOnlyField)) -and (-Not ($_.Hidden)) -and ($_.InternalName -ne "ContentType") -and ($_.InternalName -ne "Attachments") }
 
    #Loop through each item in the source and Get column values, add them to target
    [int]$Counter = 1

    #Ensureuser
        $CheckUser = Get-PnPUser -Connection $global:Con2
    #

    ForEach ($SourceItem in $SourceListItems) { 
    
       
        $ItemValue = @{}
        #Map each field from source list to target list
        Foreach ($SourceField in $SourceListFields) {
            #Check if the Field value is not Null
            If ($SourceItem[$SourceField.InternalName] -ne $Null) {
                #Handle Special Fields
                $FieldType = $SourceField.TypeAsString
                If ($FieldType -eq "User" -or $FieldType -eq "UserMulti" -or $FieldType -eq "Lookup" -or $FieldType -eq "LookupMulti") { #People Picker or Lookup Field
                    #EnsureUser
                    foreach($User in $SourceItem[$SourceField.InternalName]){
                        $Email = $User.Email
                        if($Email -ne $null){
                            
                            $CheckUser = $CheckUser | Where-Object Email -eq $Email
                            if($CheckUser -eq $null){
                                try{
                                    $NewUser = New-PnPUser -LoginName $Email
                                    $User.LookupId = $NewUser.Id
                                    #Write-Host "------------$Email Added!"
                                }catch{
                                    Write-Host "Error in copying item id = $($SourceItem.Id): $($_)" -f Red
                                }
                            }else{
                                $User.LookupId = $CheckUser.Id
                                #Write-Host "------------$Email Existed!"
                            }
                        }
                    }
                    #
                    $LookupIDs = $SourceItem[$SourceField.InternalName] | ForEach-Object { 
                        $_.LookupID.ToString() 
                    }
                    

                    if($FieldType -eq "Lookup" -or $FieldType -eq "LookupMulti"){ 
                        
                        $NewLookupIDs = @() 
                        $CheckLists = $global:Dataa | Where {$_.Id -eq $SourceField.LookupList}
                        foreach($CheckList in $CheckLists){
                            foreach($LookupID in $LookupIDs){

                                #matching old with new id

                                foreach($CheckItem in $CheckList.Items){

                                    if($CheckItem.OldId -eq $LookupID){
                                     
                                        $NewLookupIDs += $CheckItem.NewId
                                        
                                    }

                                }

                            }
                        }

                        $ItemValue.add($SourceField.InternalName, $NewLookupIDs)
                    }else{
                        $ItemValue.add($SourceField.InternalName, $LookupIDs)
                    }
                    
                }
                ElseIf ($FieldType -eq "URL") { #Hyperlink
                    $URL = $SourceItem[$SourceField.InternalName].URL
                    $Description = $SourceItem[$SourceField.InternalName].Description
                    $ItemValue.add($SourceField.InternalName, "$URL, $Description")
                }
                ElseIf ($FieldType -eq "TaxonomyFieldType" -or $FieldType -eq "TaxonomyFieldTypeMulti") { #MMS
                    $TermGUIDs = $SourceItem[$SourceField.InternalName] | ForEach-Object { $_.TermGuid.ToString() }                    
                    $ItemValue.add($SourceField.InternalName, $TermGUIDs)
                }
                Else {
                    #Get Source Field Value and add to Hashtable
                    $ItemValue.add($SourceField.InternalName, $SourceItem[$SourceField.InternalName])
                }
            }
        }
        Write-Progress -Activity "Copying List Items:" -Status "Copying Item ID '$($SourceItem.Id)' from Source List ($($Counter) of $($SourceListItemsCount))" -PercentComplete (($Counter / $SourceListItemsCount) * 100)
     
        #Copy column value from source to target
        $NewItem = Add-PnPListItem -List $TargetListName -Values $ItemValue -Connection $global:Con2

        #Copy Attachments
        Copy-SPOAttachments -SourceItem $SourceItem -TargetItem $NewItem -ListName $TargetListName 

        
        $Counter++
    }
    Write-Host "--------$($Counter - 1) items of $($SourceListItemsCount) copied"
    

}
Catch {
    Write-host -f Red "Error:" $_.Exception.Message 
}

}

Function Copy-SPOLists($Tables) {

$AllLists = $Tables

$counter = 0
foreach ($List in $AllLists) {
    $counter++
    if ($counter -lt $CopyCount) {
        Start-Sleep -s 3 
        $ListName = $List 
        $NewListName = $ListName
        
        Remove-PnPList -Identity $NewListName -Force -ErrorAction SilentlyContinue  -Connection $global:Con2
        $NewList = New-PnPList -Title $NewListName -Template GenericList -ErrorAction SilentlyContinue  -Connection $global:Con2
        Write-Host "------$NewListName List Added" -ForegroundColor Green


        #copy offlineClients
            if($SetOfflineAvailable -eq "ja"){
                $NewList.ExcludeFromOfflineClient = $false
            }else{
                $NewList.ExcludeFromOfflineClient = $true
            }
            $NewList.Update()
            Invoke-PnPQuery

        #

        try{
            $Columns = Get-PnPField -List $ListName -Connection $Con1 
        }catch{ 
            $Columns = Get-PnPField -List $ListName -Connection $Con1 #try again to catch Columns
        }

        $ColumnTitle = ""
        $ColumnType = ""
        $ColumnsList = @("Title")

        Write-Host "--------Adding Columns"
        foreach ($column in $Columns) { 
            $ColumnTitle = $column.Title
            if (($column.FromBaseType -eq $false) -and ($column.StaticName.SubString(0, 1) -ne "_")) {

                $columnType = $column.FieldTypeKind
                [Xml]$ColumnSchemaXml = $column.SchemaXml

                Remove-PnPField -List $NewListName -Identity $ColumnTitle -Force -ErrorAction SilentlyContinue  -Connection $global:Con2

                if($columnType -eq "Lookup"){

                   
                    $sourceListName = (get-pnplist -Connection $Con1 $column.LookupList).Title

                    $CheckListExist = Get-PnPList -Identity $sourceListName -ErrorAction SilentlyContinue -Connection $global:Con2

                    If($CheckListExist -eq $Null)
                    {  
                        Write-Host -f Yellow "----------$sourceListName List Does Not Exist! -> create $sourceListName list"
                        $CheckListExist = New-PnPList -Title $sourceListName -Template GenericList -ErrorAction SilentlyContinue -Connection $global:Con2
                    }

                    $CheckLookupField = Get-PnPField -List $CheckListExist -Connection $global:Con2           
                    $LookupField = $column.LookupField
                    if(($CheckLookupField | Where { $_.Title -eq $column.LookupField }).count -lt 1){
                        $LookupField = "ID"
                    }
                    

                    $NewField = Add-PnPField -List $NewListName -DisplayName $ColumnTitle -InternalName $column.InternalName -Type Lookup -Connection $global:Con2
                    Set-PnPField -List $NewListName -Identity $NewField.Id -Values  @{LookupList=(get-pnplist $sourceListName -Connection $global:Con2).Id.ToString(); LookupField=$LookupField; AllowMultipleValues= $column.AllowMultipleValues} -Connection $global:Con2                    
                }else{
					try{

                       $CorrectedSchemaXml = $ColumnSchemaXml.OuterXml.replace('&amp;',"&")
                       
                       $NewField = Add-PnPFieldFromXml -List $NewListName -FieldXml $ColumnSchemaXml.OuterXml -Connection $global:Con2

                    }catch{
                        $_
                    }
                }
                
                $ColumnsList += $ColumnTitle
                Write-Host "----------$ColumnTitle($columnType) Column added" 
                
            } 
        }

        #add Columns to "All Item"
        $InitialViews = Get-PnpView -List $NewListName -Connection $global:Con2
        
        foreach ($InitialView in $InitialViews) {
        
            $ViewColumns = $InitialView.ViewFields

            foreach($ColumnsListItem in $ColumnsList){ #check duplicate Title column in View
                if($ColumnsListItem -eq "Title" -or $ColumnsListItem -eq "LinkTitle"){
                    if(!(($ViewColumns -contains "Title") -or ($ViewColumns -contains "LinkTitle"))){
                        $ViewColumns += $ColumnsListItem
                    }
                }else{
                    $ViewColumns += $ColumnsListItem
                }
            }
            try{
                $setView = Set-PnPView -List $NewListName -Identity $InitialView.Id -Fields $ViewColumns -Connection $global:Con2    
            }
            catch{
                Write-Host "Error => $_" -ForegroundColor Red
            }
        }
        

        Write-Host "--------New Columns in Default View added"

        Copy-SPOViews -ListName $ListName -NewListName $NewListName

    }
}   
if($IsCopyPages -eq $true){
    GetAllPages  
}else{
    
    Copy-Navigation
}

}

Function Copy-SPOViews($ListName, $NewListName) {

Copy-SPOListItems -SourceList $ListName -TargetListName $NewListName

$Views = Get-PnPView -List $ListName -Includes ViewType, ViewType2, ViewFields, Aggregations, Paged, ViewQuery, RowLimit -Connection $Con1
    
#Get Properties of the source View
$Num = 0;
foreach ($View in $Views) {  
   
    $ViewProperties = @{
        "List"         = $NewListName
        "Title"        = $View.Title
        "Paged"        = $View.Paged
        "Personal"     = $View.PersonalView
        "Query"        = $View.ViewQuery
        "RowLimit"     = $View.RowLimit
        "SetAsDefault" = $View.DefaultView 
        "Fields"       = @($View.ViewFields)
        "ViewType"     = $View.ViewType
        "Aggregations" = $View.Aggregations
    }

    #Create a New View
    try{
    
                if(!$View.ServerRelativeUrl.contains("AllItems.aspx")){

                    $ExistedViews = Get-PnPView -List $NewListName -Connection $global:Con2

                    ForEach ($ExistedView in $ExistedViews) {

                        if($View.Title -eq $ExistedView.Title){
                            Remove-PnPView -List $NewListName -Identity $ExistedView.Id -Force -Connection $global:Con2
                            Write-Host "----------$($ExistedView.Title) View removed"
                        }
                    }

                    $Num++
        
                    
                    Write-Host "----------Adding '$($View.Title)' View"
                    $NewView = Add-PnPView  @ViewProperties -Connection $global:Con2
                    if($View.CustomFormatter -ne $null){
                        $View.CustomFormatter = $View.CustomFormatter.replace('< ',"&lt; ")
                    }

                    $UpdateView = Set-PnPView -List $NewListName -Identity $NewView.Id -Values @{"TabularView"=$View.TabularView; "ViewType2"=$View.ViewType2 } -Connection $global:Con2
                    
                    $UpdateView = Set-PnPView -List $NewListName -Identity $NewView.Id -Values @{"CustomFormatter" = $View.CustomFormatter} -Connection $global:Con2

                }else{
                    $ExistedViews2 = Get-PnPView -List $NewListName -Connection $global:Con2
                    ForEach ($ExistedView2 in $ExistedViews2) {
                        if($ExistedView2.ServerRelativeUrl.contains("AllItems.aspx")){

                            "--------set default view"
                            $UpdateView2= Set-PnPView -List $NewListName -Identity $ExistedView2.Id -Fields @($View.ViewFields) -Values @{"ViewQuery" = $View.ViewQuery; "RowLimit" = $View.RowLimit; "TabularView"=$View.TabularView; "ViewType2"=$View.ViewType2 } -Connection $global:Con2
                            $UpdateView2= Set-PnPView -List $NewListName -Identity $ExistedView2.Id -Fields @($View.ViewFields) -Values @{"CustomFormatter" = $View.CustomFormatter } -Connection $global:Con2

                        }
                    }
                }


    }catch{
        Write-Host "Error => $_" -ForegroundColor Red
    }
          
        
}

Write-Host "--------$Num Views added"

Write-Host "###List Copied Successfully!" -ForegroundColor Green
#Copy-SPOListItems -SourceList $ListName -TargetListName $NewListName

}

$global:TablesCurrent = @()
Function Start-Copy($Site) {

$global:Con2 = Connect-PnPOnline -Url $Site  -Interactive -ClientId $ClientId  -WarningAction Ignore

Write-Host "Connection Successfull to site: '$($Site)'" -ForegroundColor Green


If (Test-Path $FolderPath) {
    $removePath = Remove-Item -path $folderPath -recurse
}

$newPath = New-Item -Path $folderPath -ItemType Directory

$AllLists = Get-PnPList -Connection $Con1 | Where-Object { $_.Hidden -eq $false -and $_.RootFolder.ServerRelativeUrl -like "$SourcePath/lists/*" }

Write-Host "----check Dependencies" 

$Tables = @()

$counter = 0
foreach ($List in $AllLists) {

    $counter++
    if($counter -lt $CopyCount){

        $ListName = $List.Title

        if(!$List.DefaultViewUrl.contains($ListName)){
            Write-Host "Error in List $ListName => Listenname und Anzeigename der Liste m ssen gleich sein" -ForegroundColor Red
            #exit
        }

        $global:TablesCurrent = @()

        Write-Host "------check $ListName List Dependencies"
        Create-ListAndDependencies -ListName $ListName
        [array]::Reverse($global:TablesCurrent)


    
        foreach($TableCurrent in $global:TablesCurrent){
            if(!($Tables -contains $TableCurrent)){
                $Tables += $TableCurrent

                $global:Dataa += [pscustomobject]@{Title = "$TableCurrent";Id = (get-pnplist -Identity $TableCurrent -Connection $Con1).Id; Items = @()}
            }
        }
    }

}   
   

Copy-SPOLists -Tables $Tables

}

Function Create-ListAndDependencies($ListName) {

$allList = get-pnplist -Connection $Con1

$Columns = Get-PnPField -List $ListName -Connection $Con1  | Where-Object { ($_.FromBaseType -eq $false) -and ($_.StaticName.SubString(0, 1) -ne "_") -and ($_.FieldTypeKind -eq "Lookup")}

if(!($global:TablesCurrent -contains $ListName)){

    $global:TablesCurrent+= $ListName
    foreach ($column in $Columns) { 

        $SourceList = $allList | Where-Object {($_.Id -eq $column.LookupList)}
    
        Create-ListAndDependencies -ListName $SourceList[0].Title
    }
}

}

Function CopyWebparts($SourcePageName) {

$SourcePageName
  
try {        
    # $SourcePageName = Read-Host "Please enter page name from where you want to copy webparts like 'Home.aspx'"       
    if ($SourcePageName) {
        Write-Host "--------Copying $SourcePageName Page" -ForegroundColor Yellow  
        #$page = Get-PnPPage -Identity $SourcePageName -Connection $Con1 #if PnPClientSidePage not work
        $page = Get-PnPClientSidePage  -Identity $SourcePageName -Connection $Con1  
        
        $pageSections = $page.Sections
        
        $webParts = $page.Controls
        
        $WebpartsCount = $page.Controls.Count
        Write-Host "Found no. of webparts: " $WebpartsCount -ForegroundColor Gray  
        Write-Host "Copying webparts..." -ForegroundColor Gray  
        
        if ($WebpartsCount -gt 0) {

            try { 
                $DestinationName = $SourcePageName
                try {
                    Remove-PnPPage -Identity $DestinationName -Force -erroraction 'silentlycontinue' -Connection $global:Con2
                }
                catch {}
                $DestinationPage = Add-PnPPage -Name $DestinationName -Connection $global:Con2
                
                Write-Host "Adding webparts to the page: " $DestinationName -ForegroundColor Yellow  
            
                $Section = 0;
                foreach ($ps in $pageSections) {
                   
					$Section++;
                    $NewPage = Add-PnPPageSection -Page $DestinationPage -SectionTemplate $ps.Type -Order $Section -Connection $global:Con2
                   					   
					foreach ($wp in $webParts) {

                        Write-Host ($wp.SpControlData| Format-List | Out-String) -ForegroundColor Yellow

                       Write-Host ($ViewQuery| Format-List | Out-String) -ForegroundColor Yellow

						if($ps.Order -eq $wp.Section.Order){
								
							$Title = $wp.Title
							$PropertiesJson = $wp.PropertiesJson
							if (($Title -eq "Embed") -or ($Title -eq "Einbetten")) {
								$Title = "ContentEmbed"
							}
                            if (($Title -eq "Neuigkeiten")) {
								$Title = "News"
							}

                            if (($Title -eq "Websiteaktivit t")) {
								$Title = "SiteActivity"
							}
                            
                            if (($Title -eq "Dokumentbibliothek")) {
								$Title = "MyDocuments"
							}

                            if (($Title -eq "Abstandhalter")) {
								$Title = "Spacer"
							}

                            if (($Title -eq "Datei und Medien")){
                                $Title = "SpacesFileViewer"
                            }

                            
                            
							   
							if (($Title -eq "Lists") -or ($Title -eq "Liste")) {


								$Title = "List"
								$Propertiesobj = ConvertFrom-Json -InputObject $PropertiesJson  


								$ListId = $Propertiesobj.selectedListId
                                $ListName = Get-PnPList -Identity $ListId  -Connection $Con1  
                                  
                                    
                                $NewList = Get-PnPList $ListName.Title -Connection $global:Con2

                                $Propertiesobj.selectedListId = $NewList.Id

								$SelectedView = Get-PnPView -List $ListId -Identity $Propertiesobj.selectedViewId -Connection $Con1
                                
                                    
                                $TargetViews = Get-PnPView -List $NewList.Id -Connection $global:Con2
								foreach ($TargetView in $TargetViews) {                                       
									if($TargetView.Title -eq $SelectedView.Title){
                                        $Propertiesobj.selectedViewId = $TargetView.Id.Guid

                                    }                               
								}


								$PropertiesJson = ConvertTo-Json $Propertiesobj 
                               
							}                     
							
							if($wp.Type.Name -eq "PageText"){
								$NewPageTextPart = Add-PnPPageTextPart -Page $DestinationPage -Text $wp.Text -Section $Section -Column $wp.Column.Order -Connection $global:Con2
									
							}else{

								if (($wp.Title -ne $null)) {
									$NewPageTextPart = 	Add-PnPPageWebPart -Page $DestinationPage -DefaultWebPartType $Title -WebPartProperties $PropertiesJson -Section $Section -Column $wp.Column.Order -Connection $global:Con2
								}
							}
									
								
						}
					}
					   
                }
                    

                $DestinationPage.Save()  
                $DestinationPage.Publish()  

                Write-Host "Added all the webparts" -ForegroundColor Green  
            }
            catch {
                Write-Host "Error in adding webparts'" $_.Exception.Message -ForegroundColor Red               
            }
        }
        else {
            Write-Host "No webparts found'"-ForegroundColor Gray               
        }
    }
    else {
        Write-Host "Page name is empty" -ForegroundColor Red
    }
}
catch {
    Write-Host "Error in getting webparts from:'$($PageName)'" $_.Exception.Message -ForegroundColor Red               
} 

}

Function CopyWebparts2($SourcePageName) {

Write-Host  "------Copy $SourcePageName Page"
  
try {             
    if ($SourcePageName) {

        Export-PnPPage -Force -Identity $SourcePageName -Out "$LocalFolder\export.txt"  -Connection $Con1

        Invoke-PnPSiteTemplate -Path "$LocalFolder\export.txt" -WarningAction SilentlyContinue -Connection $global:Con2
       
     }
}
catch {
    Write-Host "Error in getting webparts from:'$($PageName)'" $_.Exception.Message -ForegroundColor Red               
} 

}

Function GetAllPages {

$SitePages = Get-PnPListItem -List "Site Pages" -Connection $Con1 -ErrorAction SilentlyContinue -IncludeContentType

if($SitePages -eq $null){
    $SitePages = Get-PnPListItem -List "Websiteseiten" -Connection $Con1
}

ForEach($Page in $SitePages)
{
    CopyWebparts2 -SourcePageName $Page["FileLeafRef"]
    AddImages -page $Page
}

$SelectedHomePage = Get-PnPHomePage -Connection $Con1

Set-PnPHomePage -RootFolderRelativeUrl $SelectedHomePage -Connection $global:Con2

Copy-Navigation

}

Function AddImages ($page){

$SourcePage = Get-PnPPage -Identity $Page.FieldValues.FileLeafRef -Connection $Con1
$SourceControls = $SourcePage.Controls
    

foreach($SourceControl in $SourceControls){
    if($SourceControl.Title -eq "Banner"){
         try{                       
            $ObjT = ConvertFrom-Json $SourceControl.ServerProcessedContent
            $ImageSourcesT = $ObjT.imageSources.imageSource

            if($ImageSourcesT -like "*$SourcePath*"){

                if($ImageSourcesT -like "*$SiteDomain*"){
                    $fullTargetPath = "$($SiteDomain)/$($SourcePath)"
                    $TargetImagePathT = $ImageSourcesT.Substring($fullTargetPath.Length)
                }else{
                    $TargetImagePathT = $ImageSourcesT.Substring($SourcePath.Length + 1)
                }

                $OriginalImageNameT = $TargetImagePathT.Split("/")[-1]

                Get-PnPFile -Url $ImageSourcesT -Path $LocalFolder -FileName $OriginalImageNameT -AsFile -Connection $Con1 -Force

                $TargetImagePathTT = $TargetImagePathT.substring(0,$TargetImagePathT.length - $TargetImagePathT.Split("/")[-1].Length - 1)
        
                if($TargetImagePathTT.length -gt 1){
                    $add = Add-PnPFile -Path "$($LocalFolder)\$($OriginalImageNameT)" -Folder $TargetImagePathTT -NewFileName $OriginalImageNameT -Connection $global:Con2
                }
                Remove-Item -Path "$($LocalFolder)\$($OriginalImageNameT)"
            }
        }catch{
            Write-Host  "--------Can't copy webpart Images" -ForegroundColor Red
        }

    }

    try{
        if($SourceControl.Title -eq "Bild" -OR $SourceControl.Title -eq "Image"){
            #
            $Obj = ConvertFrom-Json $SourceControl.Properties
            $ImageName = $Obj.fileName
            $OriginalImageName = $Obj.fileName
            Write-Host  "--------Copy Page Image ($($ImageName))"
            $Obj2 = ConvertFrom-Json $SourceControl.ServerProcessedContent
            $ImageSources = $Obj2.imageSources.imageSource
            
            try{
                Get-PnPFile -Url $imageSources -Path $LocalFolder -FileName $ImageName -AsFile -Connection $Con1 -Force
            }catch{
                $filePath = $Obj.advancedImageEditorData.originalSourceUrl
                $ImageName = $Obj.advancedImageEditorData.originalFileName
                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($filePath, "$($LocalFolder)\$($ImageName)")
            }

            $TargetImagePath = ""
            
            if($imageSources -like "*$SourcePath*"){
                $TargetImagePath = $imageSources.Substring($SourcePath.Length + 1)
                $OriginalImageName = $TargetImagePath.Split("/")[-1]
                $TargetImagePath = $TargetImagePath.substring(0,$TargetImagePath.length - $TargetImagePath.Split("/")[-1].Length - 1)
            }

            if($TargetImagePath.length -gt 1){
                $add = Add-PnPFile -Path "$($LocalFolder)\$($ImageName)" -Folder $TargetImagePath -NewFileName $OriginalImageName -Connection $global:Con2
            }
            Remove-Item -Path "$($LocalFolder)\$($ImageName)"
            
        }
    }catch{
        Write-Host  "--------Can't copy webpart Images" -ForegroundColor Red
    }
}

}

Function Copy-Navigation {
if($IsCopyNavigation -eq $true){

    Write-Host "------Coping Navigation Items"

    if($CopyCount -lt 2){
        Start-Sleep -Seconds 3
    }

    $SourceNavsMenu = Get-PnPNavigationNode -Location "Quicklaunch" -Connection $Con1

    $TargetNavsMenu = Get-PnPNavigationNode -Location "Quicklaunch" -Connection $global:Con2

    foreach($TargetNav in $TargetNavsMenu) { 
        Remove-PnPNavigationNode -Identity $TargetNav.Id -Force -Connection $global:Con2
    }

    foreach($SourceNav in $SourceNavsMenu) { 
 
        $parent = Add-navigation-item -SourceNav $SourceNav -TargetNavsMenu $TargetNavsMenu -parent 0 
        
        #if have children

        if($SourceNav.Children){
            $node = Get-PnPNavigationNode -Id $SourceNav.Id -Connection $Con1
            foreach($Child in $node.Children) { 
                if($parent -ne "0"){

                    $parent2 = Add-navigation-item -SourceNav $Child -TargetNavsMenu $TargetNavsMenu -parent $parent 
                    
                    #if have second child
                        $ChildNode = Get-PnPNavigationNode -Id $Child.Id -Connection $Con1
                        
                        if($ChildNode.Children -and $parent -ne "0"){
                            foreach($ChildLevel2 in $ChildNode.Children) { 
                                $sag2 = Add-navigation-item -SourceNav $ChildLevel2 -TargetNavsMenu $TargetNavsMenu -parent $parent2 -Connection $global:Con2
                            }
                        }
                    #
                }
            }
        }

      ##

    }  

}

if($IsCopyTemplateDesign -eq $true){

    $filePathToTask = "$LocalFolder\xml.txt"

    $WebSettingsTemp = Get-PnPSiteTemplate -Handlers WebSettings -Connection $Con1
    $Xml1 = [xml]$WebSettingsTemp
    $Xml1.Save($filePathToTask)


    if($CopyFromMainToSubSite -eq $true){
        $content = Get-Content -Path $filePathToTask
        $newContent = $content -replace 'RootSite', 'Web'
        $newContent | Set-Content -Path $filePathToTask
    }


    $set = Invoke-PnPSiteTemplate -Path $filePathToTask -Connection $global:Con2

    #set HeaderLayout
    $Web = Get-PnPWeb -Connection $Con1

    $WebProperty = Get-PnPProperty -ClientObject $Web -Property HeaderLayout -Connection $Con1
    if($Web.HeaderLayout){
        $HeaderLayout = Set-PnPWeb -HeaderLayout $Web.HeaderLayout -Connection $global:Con2
    }

}

if($IsCopyRegionalSettings -eq $true){
    Copy-Regionalsettings
}

}

Function Add-navigation-item($SourceNav, $TargetNavsMenu, $parent) {

$AddedId = 0

if((($SourceNav.Title -ne "Notebook") -and ($SourceNav.Title -ne "Notizbuch")) -and (($SourceNav.Title -ne "Documents") -and ($SourceNav.Title -ne "Dokumente")) -and (($SourceNav.Title -ne "Zuletzt verwendet") -and ($SourceNav.Title -ne "Recent")) ){
       #-and ($SourceNav.Title -ne "Home") 
        $IsExist = 0;
        #foreach($TargetNav in $TargetNavsMenu) {  

        #    if($TargetNav.Title -eq $SourceNav.Title){
        #        $IsExist = $TargetNav.Id;
        #    }
        #}

        if($IsExist -eq 0 -or $ReplaceTarget){
            $isFailed = $false
            $errorText = ""
           
            try{
            
                if($CopyFromList){
                    $NewUrl = $SourceNav.Url.replace([Regex]::Escape($SourcePath),[Regex]::Escape($global:CurrentSitePath))
                }else{
                    $NewUrl = $SourceNav.Url.replace([Regex]::Escape($SourcePath),[Regex]::Escape($TargetPath))
                }

                if($parent -and $parent -ne 0){
                    $NewItem = Add-PnPNavigationNode -Title $SourceNav.Title -Url $NewUrl -Location "QuickLaunch" -Parent $parent -Connection $global:Con2

                    if($NewItem -and $NewItem.Id -and $NewItem.Id -gt 0){
                        $AddedId = $NewItem.Id
                    }
                   
                    Write-Host "--------$($SourceNav.Title)   Sub-Menu Item Added!"
                }else{
                    $NewItem = Add-PnPNavigationNode -Title $SourceNav.Title -Url $NewUrl -Location "QuickLaunch" -Connection $global:Con2
                
                    if($NewItem -and $NewItem.Id -and $NewItem.Id -gt 0){
                        $AddedId = $NewItem.Id
                    }

                    Write-Host "--------$($SourceNav.Title)   Menu Item Added!"
                }

                
            }catch{
                $isFailed = $true
                $errorText =  $_
            }

            if($isFailed -eq $true -and $CopyFromList -eq $false){
                try{
            
                    $NewUrl = "$($SiteDomain)$($SourceNav.Url)"
                   
                    if($parent -and $parent -ne 0){
                        
                        $NewItem = Add-PnPNavigationNode -Title $SourceNav.Title -Url $NewUrl -Location "QuickLaunch" -Parent $parent -Connection $global:Con2
                   
                        Write-Host "--------$($SourceNav.Title)   Sub-Menu Item Added(Second Try)!"
                    }else{
                        $NewItem = Add-PnPNavigationNode -Title $SourceNav.Title -Url $NewUrl -Location "QuickLaunch" -Connection $global:Con2
                
                        if($NewItem -and $NewItem.Id -and $NewItem.Id -gt 0){
                            $AddedId = $NewItem.Id
                        }

                        Write-Host "--------$($SourceNav.Title)   Menu Item Added(Second Try)!"
                    }
                }catch{
                    $isFailed = $true
                    Write-Host $errorText -ForegroundColor Red
                }
            }elseif($isFailed -eq $true){
                Write-Host $errorText -ForegroundColor Red
            }
        }
    
    }

    return $AddedId

}

Function Copy-Regionalsettings {

$filePathSettings = "$LocalFolder\settings.txt";

try{
    $target = Get-PnPSiteTemplate -Handlers RegionalSettings -Connection $global:Con2
    $source = Get-PnPSiteTemplate -Handlers RegionalSettings -Connection $Con1

    $xml = [xml]$target
    $sourceXml = [xml]$source

    $nodes = $xml.Provisioning.Templates.ProvisioningTemplate.RegionalSettings
    $sourceNodes = $sourceXml.Provisioning.Templates.ProvisioningTemplate.RegionalSettings

    foreach($node in $nodes) {
        if($sourceNodes.ShowWeeks){
            $node.SetAttribute("ShowWeeks", $sourceNodes.ShowWeeks);
        }
        if($sourceNodes.FirstDayOfWeek){
            $node.SetAttribute("FirstDayOfWeek", $sourceNodes.FirstDayOfWeek);
        }
        if($sourceNodes.WorkDays){
            $node.SetAttribute("WorkDays", $sourceNodes.WorkDays);
        }
        if($sourceNodes.FirstWeekOfYear){
            $node.SetAttribute("FirstWeekOfYear", $sourceNodes.FirstWeekOfYear);
        }
        if($sourceNodes.WorkDayEndHour){
            $node.SetAttribute("WorkDayEndHour", $sourceNodes.WorkDayEndHour);
        }
        if($sourceNodes.WorkDayStartHour){
            $node.SetAttribute("WorkDayStartHour", $sourceNodes.WorkDayStartHour);
        }
    }  

    $xml.Save($filePathSettings);
    $set = Invoke-PnPSiteTemplate -Path $filePathSettings -ErrorAction Stop -Connection $global:Con2

    Write-Host "Settings Copied Successfully" -F Green

}catch{
    Write-Host $_ -ForegroundColor Red
}

}

Clear-Host

if($IsDevMode){
Set-Culture -CultureInfo en-US
#$ErrorActionPreference = "SilentlyContinue"
}

if($CopyFromList){
[xml]$XmlSPO = Get-Content $XmlPathList
$Sum = 0;
$ErrorSum = 0;

Write-Host "`nItem to do: $($XmlSPO.DocumentElement.Site.Length)`n" -f Cyan

foreach ($item in $XmlSPO.DocumentElement.Site){
    $global:CurrentSitePath = $item.URL.substring($SiteDomain.Length)
    Start-Copy $item.URL
}

}else{
Start-Copy $TargetSiteUrl
}