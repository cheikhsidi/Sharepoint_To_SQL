#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
   
##Variables for Processing
$SiteUrl = "https://foundationriskpartners.sharepoint.com/sites/bidash/direports/"
$ListName="FRPS_Client_map"
 
$UserName="cmoctar@foundationrp.com"
$Password ="Geologist.677"
$BatchSize = 100
  
#Setup Credentials to connect
$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))
  
Try {
    #Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Ctx.Credentials = $Cred
  
    #Get the web and List
    $Web=$Ctx.Web
    $List=$web.Lists.GetByTitle($ListName)
    $Ctx.Load($List)
    $Ctx.ExecuteQuery()
    Write-host "Total Number of Items Found in the List:"$List.ItemCount
 
    #Define CAML Query 
    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = "<View><RowLimit>$BatchSize</RowLimit></View>"
 
    Do {  
        #Get items from the list in batches
        $ListItems = $List.GetItems($Query)
        $Ctx.Load($ListItems)
        $Ctx.ExecuteQuery()
         
        #Exit from Loop if No items found
        If($ListItems.count -eq 0) { Break; }
 
        Write-host Deleting $($ListItems.count) Items from the List...
 
        #Loop through each item and delete
        ForEach($Item in $ListItems)
        {
            $List.GetItemById($Item.Id).DeleteObject()
        } 
        $Ctx.ExecuteQuery()
        # Pausing the script in millisecod
        $Ctx.RequestTimeOut = 5000*10000
    } While ($True)
 
    Write-host -f Green "All Items Deleted!"
}
Catch {
    write-host -f Red "Error Deleting List Items!" $_.Exception.Message
}


#Read more: https://www.sharepointdiary.com/2015/10/delete-all-list-items-in-sharepoint-online-using-powershell.html#ixzz6Ln9TwXPe


#Read more: https://www.sharepointdiary.com/2015/10/delete-list-in-sharepoint-online-using-powershell.html#ixzz6Ln8VYdjh

#Read more: https://www.sharepointdiary.com/2015/10/delete-list-in-sharepoint-online-using-powershell.html#ixzz6Ln4sDleT