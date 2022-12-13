   [CmdletBinding()]            
param            
(                
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
            
    [PSCredential] $PSCredential            
)            
            
 Import-Module (Join-Path -Path ([Environment]::GetEnvironmentVariable('TEMP', [EnvironmentVariableTarget]::Machine)) -ChildPath 'FIMModule.psm1') -Verbose:$false            
$Schema = New-FIMSchema            
             
$SchemaType = New-FIMSchemaType -Name 'user'            
$SchemaType | Add-FIMSchemaAttribute -Name 'saMAccountName' -Anchor -DataType 'String' -SupportedOperation ImportOnly            
$SchemaType | Add-FIMSchemaAttribute -Name 'IsEmailDomainChange' -DataType 'String' -SupportedOperation ImportExport           
$SchemaType | Add-FIMSchemaAttribute -Name 'NewEmailDomain' -DataType 'String' -SupportedOperation ImportExport           
$SchemaType | Add-FIMSchemaAttribute -Name 'mail' -DataType 'String' -SupportedOperation ImportExport           
            
$Schema.Types.Add($SchemaType)            
                     
            
$Schema  
 
 
 
