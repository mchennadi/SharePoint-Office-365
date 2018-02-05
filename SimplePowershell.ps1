#$Url=
$adminUPN="manasa.chennadi@imglobal.com"
$orgName="imgx"
$userCredential = Get-Credential -UserName $adminUPN -Message "KrithiPotty2017"
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential

$siteCollections = Get-SPOSite -Identity "https://imgx.sharepoint.com"

foreach ($siteCollection in $siteCollections)
 
{
write-host "Hi!!!!" 
$AllWebs = Get-SPOWebs($siteCollection.url)
}



$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
 
$context.Credentials = $SPOcredentials
 
$web = $context.Web
 
$context.Load($web)
 
$context.Load($web.Webs)
 
$context.load($web.lists)

$context.load($web.Navigation.QuickLaunch)

$context.ExecuteQuery()

foreach($web in $web.Webs) {



$web.Update()
$context.ExecuteQuery()

Get-SPOWebs($web.url)

}

$web.Dispose()