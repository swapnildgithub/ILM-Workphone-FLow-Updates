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


$putExportEntriesResults = new-object -Typename Microsoft.MetadirectoryServices.PutExportEntriesResults 


            
foreach($CSEntry in $CSEntries) {            
    switch ($CSEntry.ObjectType) {            
        'user' {

                 switch ($CSEntry.ObjectModificationType) {            
                'Update' { 
                            $Error.Clear() = $null
                            $UserObjectID = $CSEntry.DN.ToString().Substring('USER='.Length)
                                    
                                    $ObjectValue = New-Object System.Collections.ArrayList

                                    $CSEntry.AttributeChanges.ValueChanges | ForEach-Object {      
                                    
                                    
                                    $null = $ObjectValue.Add($_.Value)
                                  
                                    
                                    
                                    }




                                    $ea4 = $ObjectValue[0]
                                    $mail = $ObjectValue[1]
                                    $MD = "Add"
                                    
                                    switch ($MD) {            
                                    'Add'   {            
                                                  
                                                  try{
                                                       
                                                       <#
                                                       Extract the value of mail
                                                        Enable-RemoteMailbox <objectID> -RemoteRoutingAddress <mail>
                                                       
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

 
