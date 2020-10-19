
 

#Load SharePoint CSOM Assemblies  
Add-Type -path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'  
Add-Type -path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'  
  
  
$siteUrl="https://foundationriskpartners.sharepoint.com/sites/bidash"  
$listName= "Map FRPS Department" 
$columnName="DataSource Name"
$lookupWebURL="https://foundationriskpartners.sharepoint.com/sites/bidash"  
$lookupListName="FRPS DataSource"  
  
Function RepairListLookupColumns()   
{  
     param  
    (  
        [Parameter(Mandatory=$true)] $siteUrl,  
        [Parameter(Mandatory=$true)] $listName,  
        [Parameter(Mandatory=$true)] $columnName,  
        [Parameter(Mandatory=$true)] $lookupWebURL,  
        [Parameter(Mandatory=$true)] $lookupListName  
    )  
    #Passing Credentials  
    #$credPath = 'D:\Arvind\safe\secretfile.txt'  
    #$fileCred = Import-Clixml -path $credpath  
    #$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($fileCred.UserName, $fileCred.Password)  
      
    #Get site context  
   # $siteCtx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)  
   # $siteCtx.Credentials = $Cred  
    
    #Passing Credentials  
    $userId = "cmoctar@foundationrp.com"  
    $pwd1 = Read-Host -Prompt "Enter password" -AsSecureString  
    $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd1)  
    $siteCtx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
    $siteCtx.credentials = $creds  

    #Get Loookup context  
    $lookupCtx = New-Object Microsoft.SharePoint.Client.ClientContext($lookupWebURL)  
    $lookupCtx.Credentials = $Cred  
  
    #Get web, list and column with problems  
    $siteWeb = $siteCtx.Web  
    $siteList = $siteWeb.Lists.GetByTitle($listName)  
    $siteColumn = $siteList.Fields.GetByInternalNameOrTitle($columnName)  
    $siteCtx.Load($siteWeb)  
    $siteCtx.Load($siteList)  
    $siteCtx.Load($siteColumn)  
    $siteCtx.ExecuteQuery()  
  
    #Get web and lookuplist that is source of the lookup column  
    $lookupWeb = $lookupCtx.Web  
    $lookupList = $lookupWeb.Lists.GetByTitle($lookupListName)  
    $lookupCtx.Load($lookupWeb)  
    $lookupCtx.Load($lookupList)  
    $lookupCtx.ExecuteQuery()  
  
    # Prepare the IDs that are going to be updated  
    $newLookupListID = $lookupList.ID.ToString()          
    $newLookupWebID = $lookupWeb.ID.ToString()  
      
    # Replaces the XML that defined the Lookup Column  
    Write-Host $siteColumn.SchemaXml -f DarkYellow  
    $schema = $siteColumn.SchemaXml  
    [Xml]$schemaXml = $schema  
    #$schemaXml.Field.Attributes["WebId"].'#text' = $newLookupWebID  
    $requiresUpdate = $false  
    Write-Host "ID's: " $schemaXml.Field.Attributes["List"].'#text' $newLookupListID  
    if($schemaXml.Field.Attributes["List"].'#text' -ne $newLookupListID)  
    {  
        Write-Host "Found issue with List, Is:" $schemaXml.Field.Attributes["List"].'#text' "should be" $newLookupListID -f red  
        $schemaXml.Field.Attributes["List"].'#text' = $newLookupListID  
        $requiresUpdate = $true  
    }  
    if($requiresUpdate) {  
        $siteColumn.SchemaXml = $schemaXml.OuterXml  
        Write-Host "Fixing the list lookup Id to "$schemaXml.OuterXml -f DarkGray  
        $siteColumn.Update()  
        $siteCtx.ExecuteQuery()  
        Write-Host "Lookup Field has fixed...Enjoy"  
    }  
} 


RepairListLookupColumns $siteUrl $listName $columnName $lookupWebURL $lookupListName  
  