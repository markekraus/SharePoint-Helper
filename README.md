# SharePoint-Helper
PowerShell wraper commandlets for SharePoint

Must be run on a machine where the sharepoint libraries are installed
https://www.microsoft.com/en-us/download/details.aspx?id=35588 

```powershell
# Credentials used to access the SharePoint site
$SPCRED = Get-Credential
# SharePoint Site URL
$SharePointSite = "https://mytenant.sharepoint.com/teams/IT"
# SharePoint List Name or GUID identity
$Identity = "My SharePoint List "
# Hashtable for a new Item. The Key's are the StaticName for the columns which can be obtained with Get-SharePointListColumns
$NewItem = @{
    Title = "Test Entry"
    Column3 = "This is in column 3!"
    Column4 = "This is in column 4!"
}

# Get all lists on the SharePoint Site
Get-SharePointLists -SharePointSite $SharePointSite -Credential $SPCRED  | ft title,id,identity
# Get All items form a SharePoint List
Get-SharePointListItems -SharePointSite $SharePointSite -Credential $SPCRED -Identity $Identity 
# Get Information about all Columns in a SharePoint List
Get-SharePointListColumns -Credential $SPCRED -SharePointSite $SharePointSite -Identity $Identity | ft Title,StaticName,TypeDisplayName
# Add a new item to a SharePoint list
Add-SharepointListItem -SharePointSite $SharePointSite -Credential $SPCRED -Identity $Identity -Item $NewItem
```
