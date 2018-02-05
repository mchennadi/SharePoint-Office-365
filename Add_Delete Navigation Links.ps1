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
$Url = "https://imgx.sharepoint.com/EPMO/ProjectServerArchive/"


$context = Get-SPOContext -Url $Url -UserName $UserName -Password $Password

$web = $context.Web
 
$context.Load($web)
 
$context.Load($web.Webs)

 $context.ExecuteQuery();

 $filePath = create-outputfile
  
 add-content -value "
  
 </body>
  
 </html>
  
 " -path $filePath

 foreach($web in $web.Webs) {

 add-content -value "<tr><td><span style='background-color:cyan'>$($web.Title)</td></tr>" -path $filePath

 $Nav=$web.Navigation.QuickLaunch

 $context.Load($Nav)

 $context.ExecuteQuery();

 # Adding Project Details Navigation Link
 $newNavNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
 $newNavNode.Title = "Project Details"
 $newNavNode.Url = "/EPMO/ProjectServerArchive/Project Details Pages  Project Server/$($web.Title).pdf"
 $newNavNode.AsLastNode = $true
 $context.Load($Nav.Add($newNavNode))
 $context.ExecuteQuery()
 
 Write-Host "New Project Details Navigatin Link Added"  
 
 add-content -value "<tr><td><span style='background-color:white'>Updated Project Details Url from /EPMO/ProjectServerArchive/ProjectDrilldown.aspx to $($newNavNode.Url)</td></tr>" -path $filePath

 #Deleting all unused navigation links
 for($i=$Nav.Count-1;$i -ge 0; $i--)    {

$NavTitle=$Nav[$i].Title

    if($NavTitle -eq "PWA Home Page")
    {       
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Tasks")
    {        
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Risk Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Team Contacts")
    {        
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Q&A Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Decision Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Change Mgmt Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Action Item Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Issue Log")
    {
        $Nav[$i].deleteObject();
        $context.ExecuteQuery();
        add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
    }

    if($NavTitle -eq "Project Details")
    {
        $NavUrl=$Nav[$i].Url 
            
        if($NavUrl -eq "http://sharepoint.stackexchange.com")
        {             
            $Nav[$i].deleteObject();
            $context.ExecuteQuery();
            add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath
        }

        if($NavUrl -contains "/EPMO/ProjectServerArchive/ProjectDrilldown.aspx")
        {   
            $Nav[$i].deleteObject();
            $context.ExecuteQuery();
            add-content -value "<tr><td><span style='background-color:white'>Removed $($NavTitle)</td></tr>" -path $filePath          

        }       
    }     
   }    
}
$context.Dispose()