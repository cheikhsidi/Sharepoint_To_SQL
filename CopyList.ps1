#sharepoint CSOM library
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"  
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"  
  
$siteURL = "https://foundationriskpartners.sharepoint.com/sites/Technology-Data"  
$userId = "cmoctar@foundationrp.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds  

$siteURL1 = "https://foundationriskpartners.sharepoint.com/sites/bidash"  
# $userId1 = "cmoctar@foundationrp.com"  
# $pwd1 = Read-Host -Prompt "Enter password" -AsSecureString  
$creds1 = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx1 = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL1)  
$ctx1.credentials = $creds1  

#Read more: https://www.sharepointdiary.com/2015/08/sharepoint-online-get-all-lists-using-powershell.html#ixzz6N19vRRfQ

$list1 = $ctx.Web.Lists.GetByTitle("Physical Locations")  
$list2 = $ctx1.Web.Lists.GetByTitle("Loc Physical Locations")  
$list1Items = $list1.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
$fields = $list1.Fields  
$ctx.Load($list1Items)  
$ctx.Load($list1)  
$ctx1.Load($list2)  
$ctx.Load($fields)  
$ctx.ExecuteQuery()  

foreach($item in $list1Items)
{      
Write-Host $item. ID     
$listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation      
$list2Item = $list2.AddItem($listItemInfo)       
foreach($field in $fields) 
{          
# Write-Host $field.InternalName " - " $field.ReadOnlyField   
if((-Not ($field.ReadOnlyField)) -and (-Not ($field.Hidden)) -and ($field.InternalName -ne  "Attachments") -and ($field.InternalName -ne  "ContentType"))          
{              
Write-Host $field.InternalName " - " $item[$field.InternalName]              
$list2Item[$field.InternalName] = $item[$field.InternalName]              
$list2Item.update()          
}      }  }  
$ctx.ExecuteQuery()


# CopyList $source -"Physical Locations" $destination -"Loc Physical Locations"
# #Exclude system lists
# $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
# "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
# ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library"
# "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")

# Try {
 
#     #sharepoint online powershell get all lists
#     $Lists = $ctx1.web.Lists
#     $ctx1.Load($Lists)
#     $ctx1.ExecuteQuery()
 
#     #Iterate through each list in a site  
#     ForEach($List in $Lists)
#     {
#          #Exclude System Lists
#         If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
#         {
#             #Get the List Name
#             Write-host $List.Title
#             CopyList $source -$List.Title $destination -$List.Title
#         }
#     }
# }
# catch {
#     write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
# }
# CopyList $source -"Physical Locations" $destination -"Loc Physical Locations"