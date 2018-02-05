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
 

 
#search for webs
 
        if($siteCollection.Title -eq "International Medical Group")
        {
           add-content -value "<tr style='background-color:cyan'><td>$($siteCollection.url)</td><td>TopSite</td><td>$($sitecollection.template)</td><td></td></tr>" -path $filePath
           write-host "Info: Found Title:$($siteCollection.Title) | URL: $($siteCollection.url)" -foregroundcolor green

           $AllWebs = Get-SPOWebs($siteCollection.url)
        }
 
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
 
try{
 
$context.ExecuteQuery()
 
#loop through all lists in the web
 
foreach($list in $web.lists){

 if($($web.url) -like "*ProjectServerArchive")
   {
     if($($list.id) -eq "3ddd403f-f74f-49eb-85da-674c861be394")
     {    
   
     $list.Fields | Where-Object {$_.ListItemMenu -eq $true } | Select InternalName

              Write-Host "$($list.title)"
   #
   #       $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(100)
   #       $items=$list.GetItems($query)
   #       $context.Load($items)
   #       $context.ExecuteQuery()
   #
   #       foreach($item in $items)
   #       {
   #           Write-Host "Document Name:$($item.FileLeafRef)"
             #add-content -value "<tr><td><span style='margin-left:$($pixelslist)px'>$($item.Name)</td><td>List/library</td><td>$($item.id)</td><td></td></tr>" -path $filePath
   #       }        
     }
   }

add-content -value "<tr><td><span style='margin-left:$($pixelslist)px'>$($list.title)</td><td>List/library</td><td>$($list.id)</td><td>$($list.itemcount)</td></tr>" -path $filePath
 
}
 
#loop through all webs in the web and start again to find more webs
 
$pixelsweb = $pixelsweb + 15
 
$pixelslist = $pixelslist + 15
 
foreach($web in $web.Webs) {

   if($($web.url) -like "*ProjectServerArchive")
   {
         write-host "Info: Found Title:$($web.Title) | URL: $($web.url)" -foregroundcolor White
         add-content -value "<tr style='background-color:yellow'><td><span style='margin-left:$($pixelsweb)px'>$($web.url)</td><td>Web</td><td>$($web.webtemplate)</td><td>$($web.Title)</td></tr>" -path $filePath 
   }
 
Get-SPOWebs($web.url)
 
}
 
}
 
catch{
 
write-host "Could not find web" -foregroundcolor red
 
}
 
}
 
connectToO365