#sharepoint online powershell permissions report
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#SPO Client Object Model Context
$SiteURL = "https://foundationriskpartners.sharepoint.com/sites/bidash"
$ReportFile="C:\Users\CheikhMoctar\Documents\FRP_SQL_project\permission\group_Users.csv"  
$userId = "cmoctar@foundationrp.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$Ctx.credentials = $creds

function AddToGroup($SiteURL, $GroupName, $UserAccount){
    Try {
          
        #Get the Web and Group
        $Web = $Ctx.Web
        $Group= $Web.SiteGroups.GetByName($GroupName)
    
        #ensure user sharepoint online powershell - Resolve the User
        $User=$web.EnsureUser($UserAccount)
    
        #Add user to the group
        $Result = $Group.Users.AddUser($User)
        $Ctx.Load($Result)
        $Ctx.ExecuteQuery()
    
        write-host  -f Green "User '$UserAccount' has been added to '$GroupName'"
    }
    Catch {
        write-host -f Red "Error Adding user to Group!" $_.Exception.Message
    }

}

