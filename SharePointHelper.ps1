function Validate-URL {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$True)]
        [String]$URL
    )
   Process {
        $URL | ForEach-Object {
            $CurURL = $_
            $URI = $CurURL -as [system.uri]
            if ($URI.IsAbsoluteUri -and $URI.Scheme -match 'http|https' ) {
                $true
            } else {
                $false
            }
        }
    }
} 

Function ConvertTo-HashTable {

<#
.Synopsis
Convert an object into a hashtable.
.Description
This command will take an object and create a hashtable based on its properties.
You can have the hashtable exclude some properties as well as properties that
have no value.
.Parameter Inputobject
A PowerShell object to convert to a hashtable.
.Parameter NoEmpty
Do not include object properties that have no value.
.Parameter Exclude
An array of property names to exclude from the hashtable.
.Example
PS C:\> get-process -id $pid | select name,id,handles,workingset | ConvertTo-HashTable

Name                           Value                                                      
----                           -----                                                      
WorkingSet                     418377728                                                  
Name                           powershell_ise                                             
Id                             3456                                                       
Handles                        958                                                 
.Example
PS C:\> $hash = get-service spooler | ConvertTo-Hashtable -Exclude CanStop,CanPauseandContinue -NoEmpty
PS C:\> $hash

Name                           Value                                                      
----                           -----                                                      
ServiceType                    Win32OwnProcess, InteractiveProcess                        
ServiceName                    spooler                                                    
ServiceHandle                  SafeServiceHandle                                          
DependentServices              {Fax}                                                      
ServicesDependedOn             {RPCSS, http}                                              
Name                           spooler                                                    
Status                         Running                                                    
MachineName                    .                                                          
RequiredServices               {RPCSS, http}                                              
DisplayName                    Print Spooler                                              

This created a hashtable from the Spooler service object, skipping empty 
properties and excluding CanStop and CanPauseAndContinue.
.Notes
Version:  2.0
Updated:  January 17, 2013
Author :  Jeffery Hicks (http://jdhitsolutions.com/blog)

Read PowerShell:
Learn Windows PowerShell 3 in a Month of Lunches
Learn PowerShell Toolmaking in a Month of Lunches
PowerShell in Depth: An Administrator's Guide

 "Those who forget to script are doomed to repeat their work."

.Link
http://jdhitsolutions.com/blog/2013/01/convert-powershell-object-to-hashtable-revised
.Link
About_Hash_Tables
Get-Member
.Inputs
Object
.Outputs
hashtable
#>

[cmdletbinding()]

Param(
[Parameter(Position=0,Mandatory=$True,
HelpMessage="Please specify an object",ValueFromPipeline=$True)]
[ValidateNotNullorEmpty()]
[object]$InputObject,
[switch]$NoEmpty,
[string[]]$Exclude
)

Process {
    #get type using the [Type] class because deserialized objects won't have
    #a GetType() method which is what we would normally use.

    $TypeName = [system.type]::GetTypeArray($InputObject).name
    Write-Verbose "Converting an object of type $TypeName"
    
    #get property names using Get-Member
    $names = $InputObject | Get-Member -MemberType properties | 
    Select-Object -ExpandProperty name 

    #define an empty hash table
    $hash = @{}
    
    #go through the list of names and add each property and value to the hash table
    $names | ForEach-Object {
        #only add properties that haven't been excluded
        if ($Exclude -notcontains $_) {
            #only add if -NoEmpty is not called and property has a value
            if ($NoEmpty -AND -Not ($inputobject.$_)) {
                Write-Verbose "Skipping $_ as empty"
            }
            else {
                Write-Verbose "Adding property $_"
                $hash.Add($_,$inputobject.$_)
        }
        } #if exclude notcontains
        else {
            Write-Verbose "Excluding $_"
        }
    } #foreach
        Write-Verbose "Writing the result to the pipeline"
        Write-Output $hash
 }#close process

}#end function

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
        Add-SharePointLibraries  | Out-Null
    }
    Process{
        $SharePointSite | for
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
		[Parameter(Mandatory=$true,ParameterSetName='Filter')]
        [Parameter(Mandatory=$true,ParameterSetName='NoFilter')]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true,ParameterSetName='Filter')]
        [Parameter(Mandatory=$true,ParameterSetName='NoFilter')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true,ParameterSetName='Filter')]
        [Parameter(Mandatory=$true,ParameterSetName='NoFilter')]
        [Alias("Id","Identity","Name","ListID","ListName","Title","ListTitle")]
        [String]$ListIdentity,
        [Parameter(ParameterSetName='Filter',Mandatory=$true)]
        [string[]]$FilterColumns,
        [Parameter(ParameterSetName='Filter',Mandatory=$true)]
        [AllowEmptyString()]
        [string[]]$FilterValues,
        [ValidateSet("Contains","BeginsWith","Eq","Neq","Gt","Lt","Geq","Leq","Neq","IsNotNull","IsNull")]
        [Parameter(ParameterSetName='Filter',Mandatory=$true)]
        [string[]]$FilterOperators,
        [Parameter(ParameterSetName='Filter')]
        [ValidateSet("And","Or")]
        [Alias("Bool","FilterBool")]
        [string]$FilterBoolean = "Or"
    )
    begin{
        $abort = $false
        $ParamSetName = $PsCmdLet.ParameterSetName
        Write-Verbose "ParameterSet: $ParamSetName"
        Write-Verbose "Adding Libraries."
        Add-SharePointLibraries | Out-Null
        Write-Verbose "Verifying Filters."
        if($ParamSetName -eq 'Filter' -and ($FilterColumns.Count -ne $FilterValues.Count -or $FilterColumns.Count -ne $FilterOperators.Count)){
            Write-Verbose "$("FilterColumns: {0}, FilterValues: {1}, FilterOperators: {2}" -f $FilterColumns.Count, $FilterValues.Count, $FilterOperators.Count)"
            $abort = $true
            Write-Error "Number of FilterColumns, FilterValues, and FilterOperators must match."
            return
        }
    }
    Process{
        if($Abort){Return}
        if($PSCmdlet.ShouldProcess($ListIdentity)){
            try{
                Write-Verbose "Initilizing SharePoint Site Context."
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
                Write-Verbose "Initilizing SharePoint Lists Context."
                $lists=$web.Lists
                $context.Load($lists)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error enumerting lists from $SharePointSite"
                return
            }
            Write-Verbose "Parsing List Identity."
            try{
                [system.guid]::Parse($ListIdentity)| Out-Null
                $List=$lists.GetById($ListIdentity)
            }
            catch{
                $List=$lists.GetByTitle($ListIdentity)
            }
            try{
                Write-Verbose "Initilizing SharePoint List Context for $ListIdentity."
                $context.Load($List)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Unable to find List $($ListIdentity)."
                return
            }
            try{
                $query = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(100000)
                if($ParamSetName -eq 'Filter'){
                    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
                    $CamlQuery = "<View><Query><Where>"
                    if($FilterColumns.Count -gt 1){
                        $CamlQuery += "<$FilterBoolean>"
                    }
                    for($i = 0; $i -lt $FilterColumns.Count; $i++){
                        $CamlQuery += "<{0}><FieldRef Name='{1}' />" -f $FilterOperators[$i],$FilterColumns[$i]
                        if($FilterOperators[$i] -notin "IsNotNull","IsNull"){
                            $CamlQuery += "<Value Type='Text'>{0}</Value>" -f $FilterValues[$i]
                        }
                        $CamlQuery += "</{0}>" -f $FilterOperators[$i]
                    }
                    if($FilterColumns.Count -gt 1){
                        $CamlQuery += "</$FilterBoolean>"
                    }
                    $CamlQuery += "</Where></Query></View>"
                    Write-Verbose "CAML Query: $CamlQuery"
                    $query.ViewXml = $CamlQuery
                }
                $Items = $List.GetItems($query)
                $context.Load($Items)
                $context.ExecuteQuery()
            }
            catch{
                Write-Error "Error Enumerating items from list List $ListIdentity"
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

function Add-SharePointListItem {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({Validate-URL $_})]
        [string]$SharePointSite,
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true)]
        [Alias("Identity","Id","Name","ListID","ListName","Title","ListTitle")]
        [String]$ListIdentity,
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [Alias("Items")]
        $Item
        
    )
    begin{
        Write-Verbose "Begin."
        Write-Verbose "Add Libraries."
        Add-SharePointLibraries  | Out-Null
        Write-verbose "Connecting to sharepoint site ""$SharePointSite""."
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
            $abort=$true
            return
        }
        Write-verbose "Enumerating lists."
        try{
            $lists=$web.Lists
            $context.Load($lists)
            $context.ExecuteQuery()
        }
        catch{
            Write-Error "Error enumerting lists from $SharePointSite"
            $abort=$true
            return
        }
        Write-verbose "Normalizing List identity."
        try{
            [system.guid]::Parse($ListIdentity)| Out-Null
            $List=$lists.GetById($ListIdentity)
        }
        catch{
            $List=$lists.GetByTitle($ListIdentity)
        }
        Write-verbose "Loading list context for list ""$ListIdentity""."
        try{
            $context.Load($List)
            $context.ExecuteQuery()
        }
        catch{
            Write-Error "Unable to find List $ListIdentity."
            $abort=$true
            retrun
        }
    }
    Process{
        Write-Verbose "Process"
        if($abort){return}
        $Item | foreach-object {
            $CurItem = $_
            $CurItemType = [system.type]::GetTypeArray($CurItem)
            if($CurItemType.BaseType -like "System.Object" -and $CurItemType.Name -notlike "Hashtable"){
                Write-Verbose "Item is an object."
                $ItemHash = $CurItem | ConvertTo-HashTable
            }
            elseif($CurItemType.BaseType -like "System.Object" -and $CurItemType.Name -like "Hashtable"){
                Write-Verbose "Item is a hasthable."
                $ItemHash = $CurItem
            }
            else{
                Write-Error "$($CurItemType.Name) Is not a valid item type."
                return
            }
            if($PSCmdlet.ShouldProcess("Add new item to list ""$ListIdentity"" on site ""$SharePointSite""",
                    "Add new item to list ""$ListIdentity"" on site ""$SharePointSite""?",
                    "Adding new item to list ""$ListIdentity"" on site ""$SharePointSite"".")){
                Write-Verbose "Creating New Item Hash"
                $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                $NewItem = $List.AddItem($ListItemInfo)
                $ItemHash.GetEnumerator() | ForEach-Object {
                    $NewItem[$_.Key] = $_.Value
                }
                Write-Verbose "Adding Item to list."
                try{
                    $NewItem.Update()
                    $Context.ExecuteQuery()
                    Write-verbose ($ItemHash | format-table | Out-String)
                }
                catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    Write-Error "Unable to add item. $FailedItem :  $ErrorMessage"
                }
            }
        }
    }
    end{
        Write-Verbose "End"
        if($abort){return}
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
        Add-SharePointLibraries  | Out-Null
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

<#
$SPCRED = Get-Credential
$SharePointSite = "https://mitel365.sharepoint.com/teams/IT"
$Identity = "6106130b-f217-4589-9b4a-e8b633716232"
$Identity = "BitLocker Keys"
$NewItem = @{
    Recovery_x0020_Identifier="test"
    Recovery_x0020_Identifier_x0020_="test"
    Recovery_x0020_Password="test"
    Title="test"
    Domain="test.com"
}



Get-SharePointLists -SharePointSite $SharePointSite -Credential $SPCRED  | ft title,id,identity
Get-SharePointLists -SharePointSite $SharePointSite -Credential $SPCRED  | Where-Object {$_.Title -like "*bitlocker*"} | ft title,id,identity
Add-SharepointListItem -SharePointSite $SharePointSite -Credential $SPCRED -Identity $Identity -Item $NewItem
Get-SharePointListItems -SharePointSite $SharePointSite -Credential $SPCRED -Identity $Identity 
Get-SharePointLists -SharePointSite $SharePointSite -Credential $SPCRED | Where-Object {$_.Title -like $Identity} | Get-SharePointListItems -Credential $SPCRED | Select-Object -First 5 | ft title,id,identity,listtitle
#>
