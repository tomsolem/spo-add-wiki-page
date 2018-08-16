
#region Settings

$credentials = "tom@tomsodev" # loaded from windows credentials
$siteCollectionUrl = "https://tomsodev.sharepoint.com/teams/work-171"
$siteCollectionRelativeUrl = "/teams/work-171"
$newPageRelativeName = "SitePages/moved.aspx"
$newSiteUrl = "https://www.db.no"

#endregion

#region Load PnP powershell module
function Get-ScriptDirectory{
 $Invocation = (Get-Variable MyInvocation -Scope 1).Value
 Split-Path $Invocation.MyCommand.Path
}

$ResourcesPath = Join-Path (Get-ScriptDirectory) "Resources"
$ModuleName = "SharePointPnPPowerShellOnline"
$path = "$($ResourcesPath)\PnP\SharePointPnPPowerShellOnline.psd1"

try{
        $Module = Get-Module -Name $ModuleName
        if(!$Module){
            Write "$($ModuleName) is not loaded in this session"            
            Import-Module -Name $path -ErrorAction Stop
        }
        else{
            Write "$($ModuleName) is already loaded in this session"
        }
    }
    catch{
        Write "Exception caught loading the module: OfficeDevPnP.PowerShell.Commands"
        Break
    }
#endregion

#region Add wiki page and set new home page

Write-Host "`tConnect to site collection $($GlobalAdminSiteUrl)"
try{
Connect-PnPOnline -Url $siteCollectionUrl -Credentials $credentials

$serverRelativePageUrl = "$($siteCollectionRelativeUrl)/$($newPageRelativeName)"
$pageContent = "This site has moved <a href='$($newSiteUrl)'>Go to new site</a>" 

try{
write-host "Adding wiki page"
Add-PnPWikiPage -ServerRelativePageUrl $serverRelativePageUrl -Content $pageContent #-ErrorAction SilentlyContinue
write-host "Page added"
}
catch 
{
write-host "in Error"
Set-PnPWikiPageContent -ServerRelativePageUrl $serverRelativePageUrl -Content $pageContent

}
$context = Get-PnPContext
$web = $context.Web
$context.Load($web)
$context.ExecuteQuery()

$folder = $web.RootFolder
$context.Load($folder)
$context.ExecuteQuery()

Write-host "setting new homepage"
$folder.WelcomePage = $newPageRelativeName
$folder.Update()
$context.ExecuteQuery()

} catch {Throw "Error connecting to site collection [$($siteCollectionUrl)]" }

Write-Host "Disconnect SharePoint Online"
Disconnect-PnPOnline

#endregion

