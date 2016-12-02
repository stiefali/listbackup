

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

Import-Module $PSScriptRoot\SharePointListImportExport.ps1

function CreateSimpleTitleOutput([object] $Item)
{
    $content = New-Object psobject
    $content |Add-Member  -NotePropertyName "title" -NotePropertyValue $Item["Title"]
    return $content
}

function CreateSimpleTitleObject([object] $TileItem, [Microsoft.Sharepoint.Client.ListItem] $newItem)
{
    $newItem["Title"] = $TileItem.title
    $newItem.update()
    $clientContext.ExecuteQuery()
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
         $outputList+= CreateSimpleTitleOutput $Item
    }

    SaveOutputFile $file $outputList
}

function ExportViewsToJson ([Microsoft.SharePoint.Client.ClientContext]$clientContext, [String]$file) {
   
    $Items = GetListItems $clientContext "Views"
    $outputList = @()

    ForEach($Item in $Items) 
    {
        $outputList+= CreateSimpleTitleOutput $Item
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
        $outputList+= CreateSimpleTitleOutput $Item
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
     CreateSimpleTitleObject $TileItem $newItem
}

function ImportViewItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
     CreateSimpleTitleObject $TileItem $newItem
}
function ImportConfigurationItem([Microsoft.Sharepoint.Client.ListItem] $newItem, [object] $TileItem)
{
     CreateSimpleTitleObject $TileItem $newItem
}

function ImportTilesFromJson([String]$file)
{
    ImportListFromJson  $clientContext $File "Tiles" "ImportTileItem"
}
function ImportTagsFromJson([String]$file)
{
    ImportListFromJson  $clientContext $File "Tags" "ImportTagItem"
}
function ImportViewsFromJson([String]$file)
{
    ImportListFromJson  $clientContext $File "Views" "ImportViewItem"
}
function ImportTileVariantsFromJson([String]$file)
{
    ImportListFromJson  $clientContext $File "TileVariants" "ImportTileVariantItem"
}
function ImportThemesFromJson([String]$file)
{
    ImportListFromJson  $clientContext $File "Themes" "ImportThemeItem"
}
function ImportConfigurationFromJson ([String]$file)
{
    ImportListFromJson  $clientContext $File "Configuration" "ImportConfigurationItem"
}


$clientContext = CreateSharePointClientContext $SiteURL $cred

#ExportThemesToJson $clientContext $File
ImportThemesFromJson $File