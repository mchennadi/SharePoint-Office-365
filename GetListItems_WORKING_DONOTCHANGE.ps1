[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")


 

 
function create-outputfile(){

$date = get-date -format dMMyyyyhhmm
$filePath = "$($PSScriptRoot)\Output$($date).html" 
if (!(Test-Path -path $filePath)){
New-Item $filePath -type file | out-null
write-host “File created: $($filePath)” -foregroundcolor green

add-content -value "
 
<html>
 
<body>
 
<h1>Sites information Office 365</h1>
 
<table border='1' style='font-family: Calibri, sans-serif'>
 
<tr>
 
<th style='background-color:blue; color:white'>ID</th>
 
<th style='background-color:blue; color:white'>File Name</th>
 
<th style='background-color:blue; color:white'>File URL</th> 

 
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

Function Get-SPOContext([string]$Url,[string]$UserName,[string]$Password)
{
    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)    
    return $context
}

Function Get-ListItems([Microsoft.SharePoint.Client.ClientContext]$Context, [String]$ListTitle) {
    $list = $Context.Web.Lists.GetByTitle($listTitle)
    $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $items = $list.GetItems($qry)
    $Context.Load($items)
    $Context.ExecuteQuery()
    return $items 
}



$UserName = "manasa.chennadi@imglobal.com"
$Password = Read-Host -Prompt "Enter the password"    
$Url = "https://imgx.sharepoint.com/EPMO/ProjectServerArchive/"


$context = Get-SPOContext -Url $Url -UserName $UserName -Password $Password
$items = Get-ListItems -Context $context -ListTitle "Project Details Pages - Project Server" 

 $filePath = create-outputfile
  
 add-content -value "
  
 </body>
  
 </html>
  
 " -path $filePath


for($i=0;$i -lt $items.Count;$i++)
{
  Write-Host "Hi $($items[$i].Id) --> $($items[$i]['FileLeafRef']) --> /EPMO/ProjectServerArchive/Project Details Pages  Project Server/$($items[$i]['FileLeafRef'])  "

  add-content -value "<tr><td><span style='margin-left:$($pixelslist)px'>$($items[$i].Id)</td><td>$($items[$i]['FileLeafRef'])</td><td>/EPMO/ProjectServerArchive/Project Details Pages&nbsp; Project Server/$($items[$i]['FileLeafRef'])</td></tr>" -path $filePath
  
}



$context.Dispose()