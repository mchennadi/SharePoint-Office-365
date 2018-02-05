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
 
<th style='background-color:white; color:white'>Status</th>
 
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
$Url = "https://imgx.sharepoint.com/EPMO/ProjectServerArchive/PS0181 - Mission Plus Class Code/"


$context = Get-SPOContext -Url $Url -UserName $UserName -Password $Password

$web = $context.Web
 
$context.Load($web)
 
#$context.Load($web.Webs)

 $context.ExecuteQuery();

 $filePath = create-outputfile
  
 add-content -value "
  
 </body>
  
 </html>
  
 " -path $filePath

 #foreach($web in $web.Webs) {

 #add-content -value "<tr><td><span style='background-color:cyan'>$($web.Title)</td></tr>" -path $filePath

 $Nav=$context.Web.Navigation.QuickLaunch

 $context.Load($Nav)

 $context.ExecuteQuery();
 

 #Deleting all unused navigation links
 for($i=$Nav.Count-1;$i -ge 0; $i--)    {

    $NavTitle=$Nav[$i].Title

    if(!($NavTitle -eq "Site contents"   -or
        $NavTitle -eq "Project Details"  -or
        $NavTitle -eq "Documents"        -or
        $NavTitle -eq "Home"             -or
        $NavTitle -eq "Recent"))
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }
    
   }    
 
#}
$context.Dispose()