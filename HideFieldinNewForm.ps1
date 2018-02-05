Add-PSSnapin Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
#Get the Web
$SPWeb = Get-SPOObject -Username manasa.chennadi@imglobal.com -password KrithiPotty2017 -url https://imgx.sharepoint.com/Implementation -object "web/lists/getbytitle('Active Implementations')"
#Get the List
#$SPList = $SPWeb.Lists["Active Implementations"]
#Get the Field
#$SPField = $SPList.Fields["Implementation ID"]
 
#Hide from NewForm & EditForm
#$SPField.ShowInEditForm = $false
$SPField.ShowInNewForm  = $false
 
#$SPField.Update()
$SPWeb.Dispose()