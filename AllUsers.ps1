#sharepoint online powershell permissions report
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#SPO Client Object Model Context
$siteURL = "https://foundationriskpartners.sharepoint.com/sites/bidash"
# $ReportFile="C:\Users\CheikhMoctar\Documents\FRP_SQL_project\permission\group_Users.csv"  
$userId = "cmoctar@foundationrp.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds     

#Get all users of the site collection
$Users = $ctx.Web.SiteUsers
$ctx.Load($Users)
$ctx.ExecuteQuery()
 
#Get User name and Email
$Users | ForEach-Object { Write-Host "$($_.Title) - $($_.Email)"}


#Read more: https://www.sharepointdiary.com/2017/02/sharepoint-online-get-all-users-using-powershell.html#ixzz6MuX5FDrl