


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



