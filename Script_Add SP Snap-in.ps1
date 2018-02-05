if (!(test-path $profile )) 
{ 
    new-item -type file -path $profile -force 
} 
 

$cmd = 'if((Get-PSSnapin | Where-Object {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null) 
{ 
    Add-PSSnapIn "Microsoft.SharePoint.Powershell" 
}'

out-file -FilePath $profile -InputObject $cmd -Append