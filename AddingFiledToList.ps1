Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    
$siteURL = "https://foundationriskpartners.sharepoint.com/sites/bidash"  
$userId = "cmoctar@foundationrp.com"  
$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds 

#Get the Web
$Web = $Ctx.Web
$Ctx.Load($Web)
$Ctx.ExecuteQuery()

#Get All Lists from the web
$Lists = $Web.Lists
$Ctx.Load($Lists)
$Ctx.ExecuteQuery()

#Exclude system lists
$ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
"Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library"
"Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
    
$Counter = 0
#Get all lists from the web   
ForEach($List in $Lists)
{
    #Exclude System Lists
    If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
    {
        $Counter++

        try{   
            $views = $List.views  
            $view = $views.GetByTitle("View1")  
            $viewFields = $view.ViewFields  
            $viewFields.Add("ID")      
            $view.Update()  
            $ctx.executeQuery()      
        }  
        catch{  
            write-host "$($_.Exception.Message)" -foregroundcolor red  
        } 
    }      
}
# #region ***Parameters***
# $SiteURL="https://foundationriskpartners.sharepoint.com/sites/bidash/direports/"
# $ReportFile="C:\Users\CheikhMoctar\Documents\FRP_SQL_project\permission\SitePermissionRpt2.csv"
# $BatchSize = 500
# #endregion
 
# #Call the function to generate permission report
# Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile
# #Generate-SPOSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -ScanItemLevel -IncludeInheritedPermissions


#Read more: https://www.sharepointdiary.com/2018/09/sharepoint-online-site-collection-permission-report-using-powershell.html#ixzz6M9kIC3Wt