PARAM
(
[string] $SiteURL = "http://share.dev.local/sites/gertraud/pmOnecMOREShare",
[string] $File = "C:\development\test.txt",
[System.Management.Automation.PSCredential] $cred = [System.Net.NetworkCredential]::("gtotter", (ConvertTo-SecureString "PASSWORD" -AsPlainText -force))
)


try
{
    #Add-Type -Path ($PSScriptRoot + "\Microsoft.SharePoint.Client.dll")
    #Add-Type -Path ($PSScriptRoot + "\Microsoft.SharePoint.Client.Runtime.dll")
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
}
catch
{
    # type already installed
}


function CreateSharePointClientContext([String]$siteURL, $credentials) {
    #Bind to site collection
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
    #Load Cretentials
    if ($credentials)
    {
        $context.Credentials = $credentials
    }

    return $context;
}

function GetListItems([Microsoft.SharePoint.Client.ClientContext]$clientContext, [string] $listName)
{
    $List = $clientContext.Web.Lists.GetByTitle($listName)
    $clientContext.Load($List)
    $qry = New-Object Microsoft.SharePoint.Client.CamlQuery
    $qry.ViewXml = "<View Scope='RecursiveAll'>" +
                              "<Query>" + 
                               "</Query>" + 
                            "</View>"

    $Items = $List.GetItems($qry)
    $clientContext.Load($Items)
    $clientContext.ExecuteQuery()
    return $Items
}

function SaveOutputFile([String]$file, [Array] $outputList)
{
    if (-Not (Test-Path $file))
    {    
         new-item -path $File
    }
      
    $outputList | ConvertTo-Json | out-file "$file"
}

function ExportTilesToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Tiles"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content =  $Item["TileAsJSON"] | ConvertFrom-Json
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
        $content |Add-Member  -NotePropertyName "ordinal" -NotePropertyValue $Item["Ordinal"]

        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ExportTagsToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Tags"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content = New-Object psobject
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ExportViewsToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Views"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content = New-Object psobject
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ExportTileVariantsToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "TileVariants"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content =  $Item["TileVariantsAsJSON"] | ConvertFrom-Json
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]

        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ExportThemesToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Themes"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content =  $Item["ThemeAsJSON"] | ConvertFrom-Json
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
        $content |Add-Member  -NotePropertyName "tile" -NotePropertyValue $Item["Tile"]

        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ExportConfigurationToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Configuration"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $content = New-Object psobject
        $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
        $outputList+= $content
    }

    SaveOutputFile $file $outputList
}

function ImportTileItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem["Ordinal"] = $TileItem.ordinal;
        $newItem["TileAsJSON"] =  $TileItem | Select-Object * -ExcludeProperty title, ordinal | ConvertTo-Json -Compress
        $newItem.update()
        $clientContext.ExecuteQuery()
}

function ImportTileVariantItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem["TileVariantsAsJSON"] =  $TileItem | Select-Object * -ExcludeProperty title | ConvertTo-Json -Compress
        $newItem.update()
        $clientContext.ExecuteQuery()
}
function ImportThemeItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem["ThemeAsJSON"] =  $TileItem | Select-Object * -ExcludeProperty title | ConvertTo-Json -Compress
        $newItem.update()
        $clientContext.ExecuteQuery()
}

function ImportTagItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem.update()
        $clientContext.ExecuteQuery()
}

function ImportViewItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem.update()
        $clientContext.ExecuteQuery()
}
function ImportConfigurationItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
        $newItem["Title"] = $TileItem.title
        $newItem.update()
        $clientContext.ExecuteQuery()
}

function ImportListFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file, [String] $listName, [String] $importFunctionName)
{
    if (Test-Path $file)
    {    
        $input = Get-Content $file | ConvertFrom-Json

        $List = $clientContext.Web.Lists.GetByTitle($listName)
        $clientContext.Load($List)

        
        ForEach($TileItem in $input) 
        {
            $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation; 
            $newItem  = $List.AddItem($ListItemCreationInformation);
            & $importFunctionName $newItem $TileItem
        }

    }
    else
    {
        throw "File not found"
    }    
}

function ImportTilesFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "Tiles" "ImportTileItem"
}
function ImportTagsFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "Tags" "ImportTagItem"
}
function ImportViewsFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "Views" "ImportViewItem"
}
function ImportTileVariantsFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "TileVariants" "ImportTileVariantItem"
}
function ImportThemesFromJson([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "Themes" "ImportThemeItem"
}
function ImportConfigurationFromJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file)
{
    ImportListFromJson  $clientContext $File "Configuration" "ImportConfigurationItem"
}


$clientContext = CreateSharePointClientContext $SiteURL $cred

#ExportConfigurationToJson $clientContext $File
#ImportConfigurationFromJson $clientContext $File

