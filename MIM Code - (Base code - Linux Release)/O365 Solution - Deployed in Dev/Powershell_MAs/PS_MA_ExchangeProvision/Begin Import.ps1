    [CmdletBinding()]            
param(                
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
                
    [ValidateNotNull()]            
    [PSCredential] $PSCredential,           
            
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep] $OpenImportConnectionRunStep,            
                
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3
           
           
Import-Module (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'FIMModule.psm1') -Verbose:$false -ErrorAction Stop 
Import-Module ActiveDirectory 
 
$Server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server"

          

$MsolObjects = @(Get-ADUser -filter {msExchRecipientTypeDetails -eq 2147483648} -searchscope 2  -Properties samAccountName,mail,msExchRecipientTypeDetails |  Select-Object @{Name='Type';Expression={'User'}},mail,msExchRecipientTypeDetails,samAccountName)
           
$FilePath = Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'MsolObjects.xml'            
            
Write-Verbose "MsolObjects '$FilePath': $($MsolObjects.Count.ToString('#,##0'))"            
            
$MsolObjects | Export-Clixml -Path $FilePath -ErrorAction Stop            
            
New-FIMOpenImportConnectionResults 
 
 

 
