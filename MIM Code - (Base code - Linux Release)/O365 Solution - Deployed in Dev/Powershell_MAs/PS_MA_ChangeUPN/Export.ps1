 [CmdletBinding()]            
param            
(                
    [ValidateNotNull()]            
    [System.Collections.ObjectModel.KeyedCollection[string, Microsoft.MetadirectoryServices.ConfigParameter]] $ConfigParameters,            
            
    [PSCredential] $PSCredential,            
            
    [System.Collections.Generic.IList[Microsoft.MetaDirectoryServices.CSEntryChange]] $CSEntries,            
                
    [Microsoft.MetadirectoryServices.OpenExportConnectionRunStep] $OpenExportConnectionRunStep,            
            
    [Microsoft.MetadirectoryServices.Schema] $Schema            
)            
            
Set-StrictMode -Version 3


Enter-Script -ScriptType "Export" -ErrorObject $Error

$Group = "contoso\LegalAccess"
$Policy = "Mailbox Policy - User Mailboxes"

$putExportEntriesResults = new-object -Typename Microsoft.MetadirectoryServices.PutExportEntriesResults 


            
foreach($CSEntry in $CSEntries) {            
    switch ($CSEntry.ObjectType) {            
        'user' {

                 switch ($CSEntry.ObjectModificationType) {            
                'Update' { 
                            $Error.Clear() = $null
                            $UserObjectID = $CSEntry.DN.ToString().Substring('USER='.Length)
                                   
                                    $CSEntry.AttributeChanges.ValueChanges | ForEach-Object {            
  
                                    $ObjectValue = $_.Value
                                    switch ($_.ModificationType) {            
                                    'Add'   {            
                                                  
                                                  try{
                                                       

                                                                           <#
                                                                           -> Get the substring of mail before @
                                                                           -> build new UPN as  above substring + NewEmailDomain
                                                                           Set-MsolUserPrincipalName -UserPrincipalName <mail> -NewUserPrincipalName <new User Principal Name>
                                                                           
                                                                           
                                                                           #>


                                                                               
                                                                           
                                                                       
                                                        break
                                                     }
                                                  catch {

                                                        Write-Error "$_"

                                                  }

                                                    
                                            }

                                                 
                                    'Delete' {

 
                                                        break    

                                             }

                                    
                                    

                                    }
 
                                     
                                    }
                                          
            
                          
        
                         }
               }
           }     
       }

  
  if ($Error)
		{
			$csentryChangeResult = New-CSEntryChangeExportError -CSEntryChangeIdentifier $csentry.Identifier -ErrorObject $Error
		}
		else
		{
            $csEntryChangeresult = [Microsoft.MetadirectoryServices.CSEntryChangeresult]::Create($csentry.Identifier,$csentry.AttributeChanges,"Success")
		}

  
  
  
  
  $putExportEntriesResults.CSEntryChangeResults.Add($csEntryChangeresult)
          
 }  

Write-Output $putExportEntriesResults

Exit-Script -ScriptType "Export" -SuppressErrorCheck -ErrorObject $Error 
