  [CmdletBinding()]            
param(                
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
                
    [PSCredential] $PSCredential,            
            
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep] $OpenImportConnectionRunStep,            
                
    [ValidateNotNull()]            
    [Microsoft.MetadirectoryServices.CloseImportConnectionRunStep] $CloseImportConnectionRunStep,            
                
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3            
            
$FilePath = Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'MsolObjects.xml'            
            
if ($CloseImportConnectionRunStep.Reason -eq [Microsoft.MetadirectoryServices.CloseReason]::Normal) {            
   if (Test-Path $FilePath) {            
        Remove-Item $FilePath -Confirm:$false -Force            
    }            
}     
       
New-FIMCloseImportConnectionResults             
 
