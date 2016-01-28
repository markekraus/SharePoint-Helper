function Validate-URL {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )
   Process {
        $URI = $URL -as [system.uri]
        if ($URI.IsAbsoluteUri -and $URI.Scheme -match 'http|https' ) {
            $true
        } else {
            $false
        }
    }
} 

Function Add-SharePointLibraries {

    if("Microsoft.SharePoint.Client.ClientRuntimeContext" -as [type]){
        return $true
    }
    $SharePointClientLocations = @(
        "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI",
        "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell",
        "C:\Program Files (x86)\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell"
    )
    foreach($SharePointClientLocation in $SharePointClientLocations){
        if((Test-Path "$SharePointClientLocation\Microsoft.SharePoint.Client.dll") -and (Test-Path "$SharePointClientLocation\Microsoft.SharePoint.Client.Runtime.dll")){ break; }
    }
    try{    
        Add-Type –Path "$SharePointClientLocation\Microsoft.SharePoint.Client.dll"
        Add-Type –Path "$SharePointClientLocation\Microsoft.SharePoint.Client.Runtime.dll"
    }
    catch {
        Write-Error "Unable to load SharePoint Client Libraries"
        return $false
    }
}

function Get-SharePointLists {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential
    )
    begin{
        if(!(Add-SharePointLibraries)){
            Write-Error "Unable to load SharePoint Client Libraries"
            return $false
        }
    }
    Process{
        if($PSCmdlet.ShouldProcess($SharePointSite)){
            try{
                $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointSite)
                $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName,$Credential.Password)
                $site = $context.Site
                $web = $context.Web
                $context.Load($site)
                $context.Load($web)
                $context.ExecuteQuery()
                
            }
            catch{
                Write-Error "Error in connecting to sharepoint site $SharePointSite"
                return
            }
            try{
                $lists=$web.Lists
                $context.Load($lists)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error enumerting lists from $SharePointSite"
            }
            
            for($i = 0; $i -lt $lists.Count; $i++){
                $OutObj = New-Object psobject
                $OutObj | Add-Member -MemberType NoteProperty -Name SharePointSite -Value $SharePointSite
                $OutObj | Add-Member -MemberType NoteProperty -Name Identity -Value $lists[$i].Id
                $lists[$i] | Get-Member -MemberType Property | where-object {$_.Definition -inotmatch "sharepoint" } |ForEach-Object {
                    $Name = $_.Name
                    $OutObj | Add-Member -MemberType NoteProperty -Name $Name -Value $lists[$i].$Name
                }
                $OutObj
            }
        }
    }
}

function Get-SharePointListItems {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Medium")]
	param(
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True)]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias("Id","Name","ListID","ListName","Title","ListTitle")]
        [String]$Identity
    )
    begin{
        if(!(Add-SharePointLibraries)){
            Write-Error "Unable to load SharePoint Client Libraries"
            return $false
        }
    }
    Process{
        if($PSCmdlet.ShouldProcess($Identity)){
            try{
                $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointSite)
                $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName,$Credential.Password)
                $site = $context.Site
                $web = $context.Web
                $context.Load($site)
                $context.Load($web)
                $context.ExecuteQuery()
                
            }
            catch{
                Write-Error "Error in connecting to sharepoint site $SharePointSite"
                return
            }
            try{
                $lists=$web.Lists
                $context.Load($lists)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error enumerting lists from $SharePointSite"
                return
            }
            try{
                [system.guid]::Parse($Identity)| Out-Null
                $List=$lists.GetById($Identity)
            }
            catch{
                $List=$lists.GetByTitle($Identity)
            }
            try{
                $context.Load($List)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Unable to find List $($Identity)."
                return
            }
            try{
                $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(100000)
                $Items = $List.GetItems($query)
                $context.Load($Items)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error Enumerating items from list List $Identity"
                return
            }
                 
            for($i = 0; $i -lt $Items.Count; $i++){
                $OutObj = New-Object psobject
                $OutObj | Add-Member -MemberType NoteProperty -Name SharePointSite -Value $SharePointSite
                $OutObj | Add-Member -MemberType NoteProperty -Name ListIdentity -Value $List.Id
                $OutObj | Add-Member -MemberType NoteProperty -Name ListTitle -Value $List.Title
                $OutObj | Add-Member -MemberType NoteProperty -Name Identity -Value $Items[$i].ID
                $Items[$i].FieldValues.Keys |ForEach-Object {
                    $Name = $_
                    $Value = $Items[$i].FieldValues.$_
                    $OutObj | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
                }
                $OutObj
            }
        }
    }
}

function Add-SharepointListItem {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Medium")]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias("Id","Name","ListID","ListName","Title","ListTitle")]
        [String]$Identity,
        [hashtable]$Item
    )
    begin{
        if(!(Add-SharePointLibraries)){
            Write-Error "Unable to load SharePoint Client Libraries"
            return $false
        }
    }
    Process{
        if($PSCmdlet.ShouldProcess($Item)){
            try{
                $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointSite)
                $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName,$Credential.Password)
                $site = $context.Site
                $web = $context.Web
                $context.Load($site)
                $context.Load($web)
                $context.ExecuteQuery()
                
            }
            catch{
                Write-Error "Error in connecting to sharepoint site $SharePointSite"
                return
            }
            try{
                $lists=$web.Lists
                $context.Load($lists)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error enumerting lists from $SharePointSite"
                return
            }
            try{
                [system.guid]::Parse($Identity)| Out-Null
                $List=$lists.GetById($Identity)
            }
            catch{
                $List=$lists.GetByTitle($Identity)
            }
            try{
                $context.Load($List)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Unable to find List $Identity."
                return
            }
            $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $NewItem = $List.AddItem($ListItemInfo)
            foreach($Key in $Item.Keys){
                $NewItem[$Key] = $Item.$key
            }
            try{
                $NewItem.Update()
                $Context.ExecuteQuery()
            }
            catch{
                Write-Error "Unable to add item`r`n$($NewItem | out-string -stream)"
            }
        }
    }
}

function Get-SharePointListColumns {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Medium")]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity,
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential
        
    )
    begin{
        if(!(Add-SharePointLibraries)){
            Write-Error "Unable to load SharePoint Client Libraries"
            return $false
        }
    }
    Process{
        foreach($CurIdentity in $Identity) {
            if($PSCmdlet.ShouldProcess($CurIdentity)){
                try{
                    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointSite)
                    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName,$Credential.Password)
                    $site = $context.Site
                    $web = $context.Web
                    $context.Load($site)
                    $context.Load($web)
                    $context.ExecuteQuery()
                    
                }
                catch{
                    Write-Error "Error in connecting to sharepoint site $SharePointSite"
                    return
                }
                try{
                    $lists=$web.Lists
                    $context.Load($lists)
                    $context.ExecuteQuery()
                }
                catch{
                    Write-Error "Error enumerting lists from $SharePointSite"
                    return
                }
                try{
                    [system.guid]::Parse($CurIdentity)| Out-Null
                    $List=$lists.GetById($CurIdentity)
                }
                catch{
                    $List=$lists.GetByTitle($CurIdentity)
                }
                try{
                    $context.Load($List)
                    $context.ExecuteQuery()
                }
                catch{
                    Write-Error "Unable to find List $($CurIdentity)."
                    return
                }
                try{
                    $Fields = $List.Fields
                    $context.Load($Fields)
                    $context.ExecuteQuery()
                }
                catch{
                    Write-Error "Error Enumerating columns from list List $Identity"
                    return
                }
                     
                $Fields | Select-Object @{
                    Name="SharePointSite";
                    Expression={$SharePointSite}},@{
                    Name="ListIdentity";
                    Expression={$List.Id}},@{
                    Name="ListTitle";
                    Expression={$List.Title}},@{
                    Name="Identity";
                    Expression={$_.Id}},*
            }
        }
    }
}
