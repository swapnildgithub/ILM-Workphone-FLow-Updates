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

$Policy = "Mailbox Policy - User Mailboxes"
$DefaultPolicy="Default MRM Policy"

$putExportEntriesResults = new-object -Typename Microsoft.MetadirectoryServices.PutExportEntriesResults 

$rulesfile = Join-Path -Path ("C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Extensions") -ChildPath 'rules-config.xml'             

$RulesConfig =  [xml] (get-content $rulesfile)

$PAC_CODES = @($RulesConfig.SelectNodes("/rules-extension-properties/management-agents/dev/PACCodeForUSandPR"))

$PSCredential >> C:\logs\abc.txt
            
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
                                                        
                                                        $Result =""
                                                        $Result = $PAC_CODES | where {$_.CODE -eq  $ObjectValue}

                                                        
                                                                IF([string]::IsNullOrWhiteSpace($Result)) 
                        
                                                                     { 

                                                                            
                                                                            
                                                                            Set-Mailbox -identity $UserObjectID -ManagedFolderMailboxPolicy $null -ErrorAction Stop
                                                                            
                                                                     } else { 
                                                                              
                                                                             Set-Mailbox -identity $UserObjectID -ManagedFolderMailboxPolicy $Policy -ErrorAction Stop -Force

                                                                             
                                                                     }
                                                                    Set-ADUser -Identity $UserObjectID -Clear extensionAttribute4 -Credential $PSCredential 
                                                                    Set-ADUser -Identity $UserObjectID -Add @{extensionAttribute4=$ObjectValue} -Credential $PSCredential 
                                                                    Set-CASMailbox $UserObjectID -PopEnabled $false -ImapEnabled $false -DomainController "$domainController" -ErrorAction Stop
                                                                       
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
 
 
 
