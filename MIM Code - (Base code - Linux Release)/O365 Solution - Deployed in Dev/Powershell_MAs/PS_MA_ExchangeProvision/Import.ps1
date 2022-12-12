      [CmdletBinding()]            
param(                
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
                
    [PSCredential] $PSCredential,            
            
    [ValidateNotNull()]            
    [Microsoft.MetadirectoryServices.ImportRunStep] $GetImportEntriesRunStep,            
            
    [ValidateNotNull()]            
    [Microsoft.MetadirectoryServices.OpenImportConnectionRunStep] $OpenImportConnectionRunStep,            
                
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3

Enter-Script -ScriptType "Import" -ErrorObject $Error
            
$FilePath = Join-Path -Path ([Microsoft.MetadirectoryServices.MAUtils]::MAFolder) -ChildPath 'MsolObjects.xml'            
            
if (!(Test-Path $FilePath)) {            
    throw "Could not locate import file: $FilePath"            
}            
            
if ($GetImportEntriesRunStep.CustomData) {            
    $Skip = [int]$GetImportEntriesRunStep.CustomData            
} else {            
    $Skip = 0            
}            
          
$MsolObjects = @(Import-Clixml -Path $FilePath -Skip $Skip -First $OpenImportConnectionRunStep.PageSize -ErrorAction Stop)            
            
$Changes = New-FIMCSEntryChanges            
            
$MsolObjects | ForEach-Object {            
    $MsolObj = $_
            
    $CSEntry = $null            
            
    switch ($MsolObj.Type) {            
        'User' {
        

                        $whereToProvision = ""
                        IF([string]::IsNullOrWhiteSpace($MsolObj.whereToProvision)) 
                        
                        { 
                          $whereToProvision = "TBD"         
                                
                        } else {            
                        
                           $whereToProvision = $MsolObj.whereToProvision          
                        }  
    
    $msExchRecipientTypeDetails=[String]$MsolObj.msExchRecipientTypeDetails
                    
            $CSEntry = New-FIMCSEntryChange -ObjectType 'user' -ModificationType Add -DN "USER=$($MsolObj.saMAccountName)" |
            Add-FIMCSAttributeChange -Add -Name 'msExchRecipientTypeDetails' -Value $msExchRecipientTypeDetails -PassThru 

             
                                                 
            }            
                
           
    }            
            
    $Changes.Add($CSEntry)            
}            
            
New-FIMGetImportEntriesResults -CustomData "$($Skip + $OpenImportConnectionRunStep.PageSize)" -MoreToImport:($MsolObjects.Count -ge $OpenImportConnectionRunStep.PageSize) -CSEntries $Changes 
 
 Exit-Script -ScriptType "Import" -SuppressErrorCheck -ErrorObject $Error 
 
 
 
