function connectToO365{
 
# Let the user fill in their admin url, username and password
 
$adminUrl = Read-Host "Enter the Admin URL of 0365 (eg. https://<Tenant Name>-admin.sharepoint.com)"
 
$userName = Read-Host "Enter the username of 0365 (eg. admin@<tenantName>.onmicrosoft.com)"
 
$password = Read-Host "Please enter the password for $($userName)" -AsSecureString
 
# set credentials
 
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $password
 
$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $password)
 
#connect to to Office 365
 
try{
 
Connect-SPOService -Url $adminUrl -Credential $credentials
 
write-host "Info: Connected succesfully to Office 365" -foregroundcolor green
 
}
 
catch{
 
write-host "Error: Could not connect to Office 365" -foregroundcolor red
 
Break connectToO365
 
}
 
#create HTML file
 
$filePath = create-outputfile
 
#start getting site collections
 
get-siteCollections
 
add-content -value "
 
</body>
 
</html>
 
" -path $filePath
 
}

function create-outputfile(){
 
#Create unique string from the date
 
$date = get-date -format dMMyyyyhhmm
 
#set the full path
 
$filePath = "$($PSScriptRoot)\Output$($date).html"
 
#test path
 
if (!(Test-Path -path $filePath)){
 
#create file
 
New-Item $filePath -type file | out-null
 
#print info
 
write-host “File created: $($filePath)” -foregroundcolor green
 
#add start HTML information to file
 
add-content -value "
 
<html>
 
<body>
 
<h1>Sites information Office 365</h1>
 
<table border='1' style='font-family: Calibri, sans-serif'>
 
<tr>
 
<th style='background-color:blue; color:white'>URL</th>
 
<th style='background-color:blue; color:white'>Type</th>
 
<th style='background-color:blue; color:white'>Template</th>
 
<th style='background-color:blue; color:white'>Item Count</th>
 
</tr>
 
" -path $filePath
 
}
 
else{
 
#break so there won't be duplicate files
 
write-host "Output file already exists, wait 1 minute" -foregroundcolor yellow
 
Break create-outputfile
 
}
 
return $filePath
 
}

function get-siteCollections{
 
#Get all site collections
 
$siteCollections = Get-SPOSite
 
#loop through all site collections
 
foreach ($siteCollection in $siteCollections)
 
{
 
#set variable for a tab in the table
 
$pixelsweb = 0
 
$pixelslist = 0
 
#add info to HTML document
 
add-content -value "<tr style='background-color:cyan'><td>$($siteCollection.url)</td><td>TopSite</td><td>$($sitecollection.template)</td><td></td></tr>" -path $filePath
 
write-host "Info: Found $($siteCollection.url)" -foregroundcolor green
 
#search for webs
 
$AllWebs = Get-SPOWebs($siteCollection.url)
 
}
 
}
 
function Get-SPOWebs($url){
 
#fill metadata information to the client context variable
 
$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
 
$context.Credentials = $SPOcredentials
 
$web = $context.Web
 
$context.Load($web)
 
$context.Load($web.Webs)
 
$context.load($web.lists)

$context.load($web.Navigation.QuickLaunch)
 
try{
 
$context.ExecuteQuery()
 
#loop through all lists in the web
 
foreach($list in $web.lists){
 
add-content -value "<tr><td><span style='margin-left:$($pixelslist)px'>$($list.title)</td><td>List/library</td><td></td><td>$($list.itemcount)</td></tr>" -path $filePath
 
}
 
#loop through all webs in the web and start again to find more subwebs
 
$pixelsweb = $pixelsweb + 15
 
$pixelslist = $pixelslist + 15

#$sitelogo="/SiteAssets/IMG%20Logo%20190px.png"
#$NewNavUrl="/EPMO/ProjectServerArchive/Project%20Details%20Pages%20%20Project%20Server/$($web.Title).pdf"
#$oldUrl="http://sharepoint.stackexchange.com"
 
foreach($web in $web.Webs) {

$web.Navigation.QuickLaunch | ForEach-Object {

Write-Host "I am here"

  #  if($_.Url -contains "EPMO/ProjectServerArchive"){
  #      $linkUrl = $_.Url
  #      Write-Host "Updating with new URL"
  #     # $_.Url = $_.Url.Replace($FindString,$ReplaceString)
  #      #$_.Update()
  #  }

}

#$newNavNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
#$newNavNode.Title = "Project Details"
#$newNavNode.Url = "http://sharepoint.stackexchange.com"
#$newNavNode.AsLastNode = $true
#$navColl.Add($newNavNode)
         
$web.Update()
$context.ExecuteQuery()
 
add-content -value "<tr style='background-color:yellow'><td><span style='margin-left:$($pixelsweb)px'>$($web.url)</td><td>Web</td><td>$($web.webtemplate)</td><td></td></tr>" -path $filePath
 
#write-host "Info: Found $($web.url)" -foregroundcolor green

Get-SPOWebs($web.url)


}

$web.Dispose()
 
}
 
catch{
 
write-host "Could not find web" -foregroundcolor red
 
}
 
}
 

 function Repair-SPLeftNavigation {
[CmdletBinding()]
Param(
[Parameter(Mandatory=$true)][System.String]$Web,
[Parameter(Mandatory=$true)][System.String]$FindString,
[Parameter(Mandatory=$true)][System.String]$ReplaceString
)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
$SPWeb = Get-SPWeb $Web
foreach ($subsite in $SPWeb.Webs) {

$subsite.Navigation.QuickLaunch | ForEach-Object {
    if($_.Url -match "^$FindString"){
        $linkUrl = $_.Url
        Write-Host "Updating $linkUrl with new URL"
        $_.Url = $_.Url.Replace($FindString,$ReplaceString)
        $_.Update()
    }

}
$subsite.Dispose()
}
$SPWeb.Dispose()
}#endFunction

connectToO365