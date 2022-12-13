   [CmdletBinding()]            
param(                
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
                
    [ValidateNotNull()]            
    [PSCredential] $PSCredential,           
            
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep] $OpenImportConnectionRunStep,            
                
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3
           
$ConfigParameters >> c:\logs\config.txt
           
Import-Module (Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'FIMModule.psm1') -Verbose:$false -ErrorAction Stop 
Import-Module ActiveDirectory 
 
$Server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server"



$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Server -Credential $PSCredential

$null = Import-PSSession $Session  

          
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Server -Credential $PSCredential
#$null = Import-PSSession $Session 
#$MsolObjects = @(get-mailbox |  Select-Object @{Name='Type';Expression={'User'}},saMAccountName,RetentionPolicy)       

$MsolObjects = @(Get-ADUser -filter {msExchRecipientTypeDetails -eq 2} -searchbase "OU=Standard Accounts,OU=Domain Accounts,DC=rfd,DC=lilly,DC=com" -searchscope 1  -Properties extensionAttribute4 |  Select-Object @{Name='Type';Expression={'User'}},saMAccountName,msExchRecipientTypeDetails, extensionAttribute4)
           
$FilePath = Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'MsolObjects.xml'            
            
Write-Verbose "MsolObjects '$FilePath': $($MsolObjects.Count.ToString('#,##0'))"            
            
$MsolObjects | Export-Clixml -Path $FilePath -ErrorAction Stop            
            
New-FIMOpenImportConnectionResults 
  
