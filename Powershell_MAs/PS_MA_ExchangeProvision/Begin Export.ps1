  [CmdletBinding()]            
param(                
    [ValidateNotNull()]            
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
                
    [ValidateNotNull()]            
    [PSCredential] $PSCredential,            
            
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep] $OpenExportConnectionRunStep,            
                
    [ValidateNotNull()]            
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3            
            
Import-Module (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'FIMModule.psm1') -Verbose:$false -ErrorAction Stop            

$Server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server"
Connect-MsolService -Credential $PSCredential
$msolSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Server -Credential $PSCredential -Authentication Basic -AllowRedirection
$null=Import-PSsession $msolSession  
 
