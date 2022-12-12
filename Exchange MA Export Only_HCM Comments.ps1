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
$Username = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "username_Global"
$ADPassword = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "password_Global" -Encrypted


$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Username,$ADPassword 

$Server = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName "Server"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Server -Credential $PSCredential
$null = Import-PSSession $Session  
Connect-MsolService -Credential $cred 
 
  
                                                                                                                [CmdletBinding()]

param(

	[parameter(Mandatory = $true)]

	[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]]

	$ConfigParameters,

	[parameter(Mandatory = $true)]

	[Microsoft.MetadirectoryServices.Schema]

	$Schema,

	[parameter(Mandatory = $true)]

	[Microsoft.MetadirectoryServices.OpenExportConnectionRunStep]

	$OpenExportConnectionRunStep,

	[parameter(Mandatory = $true)]

	[System.Collections.Generic.IList[Microsoft.MetadirectoryServices.CSEntryChange]]

	$CSEntries,

	[parameter(Mandatory = $true)]

	[Alias('PSCredential')] # To fix mess-up of the parameter name in the RTM version of the PowerShell connector.

	[System.Management.Automation.PSCredential]

	$Credential,

	[parameter(Mandatory = $false)]

	[ValidateScript({ Test-Path $_ -PathType "Container" })]

	[string]

	$ScriptDir = [Microsoft.MetadirectoryServices.MAUtils]::MAFolder # Optional parameter for manipulation by the TestHarness script.

)

#region configurable fields

$rulesfile = Join-Path -Path ("C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Extensions") -ChildPath 'rules-config.xml'             
$RulesConfig =  [xml] (get-content $rulesfile)
$environment=$RulesConfig.SelectSingleNode("/rules-extension-properties/environment").InnerText
$domainController=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+ $environment +"/DomainController").InnerText
$UPNDomain=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/UPNDomain").InnerText
$RBAC=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/Exchange/RBAC").InnerText
$parentOU=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/ad-ma/parentOU").InnerText
$logFileLocation=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/ExchangeMAlogFileLocation").InnerText
$msExchRecipientDisplayType_RemoteMailbox=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientDisplayType_RemoteMailbox").InnerText
$msExchRecipientDisplayType_LocalMailbox=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientDisplayType_LocalMailbox").InnerText
$msExchRecipientDisplayType_Contact=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientDisplayType_Contact").InnerText
$msExchRecipientTypeDetails_LocalMailbox=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientTypeDetails_LocalMailbox").InnerText
$msExchRecipientTypeDetails_LocalMailboxwithoutSID=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientTypeDetails_LocalMailboxwithoutSID").InnerText
$msExchRecipientTypeDetails_RemoteMailbox=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientTypeDetails_RemoteMailbox").InnerText
$msExchRecipientTypeDetails_ContactWithSID=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientTypeDetails_ContactWithSID").InnerText
$msExchRecipientTypeDetails_ContactWithoutSID=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/"+$environment+"/msExchRecipientTypeDetails_ContactWithoutSID").InnerText
$AMdomainController=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/AMdomainController").InnerText
$bogusdomain=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/ad-ma/bogusdomain").InnerText
$legacyExchangeDNWithoutEmpID=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/Exchange/legacyExchangeDN").InnerText
$exchangeFlag=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/Exchange/exchangeFlag").InnerText
$configpath="C:\Program Files\Microsoft Identity Integration Server\Extensions\$environment.xml"
[System.Xml.XmlDocument]$xdExchange = new-object System.Xml.XmlDocument
$xdExchange = ([xml]( Get-Content $configpath ))
$parentnodelist=$xdExchange.SelectNodes("/MIIS-AD-CONFIG/PARENT")  

#endregion




Set-StrictMode -Version "2.0"



$commonModule = (Join-Path -Path $scriptDir -ChildPath $configParameters["FIMModule.psm1"].Value)



if (!(Get-Module -Name (Get-Item $commonModule).BaseName)) { Import-Module -Name $commonModule }



Enter-Script -ScriptType "Export" -ErrorObject $Error



# HCM - This function is called from Export-User Mehtod for Update Modification type attribute checks.
function isModified($strAttribute)
{
    if($CSEntryChange.AttributeChanges[$strAttribute] -ne $null)
    {
        if($CSEntryChange.AttributeChanges[$strAttribute].ModificationType -ne $null)
        {
        return $true
        }
        else
        {
        return $false
        }

    }
    else
    {
    return $false
    }
}
# HCM - This function is called from Export-User Mehtod for Add or Update type changes.
function set_RecTypeDetails_RecDisplayType
{
[CmdletBinding()]
	param(
		[parameter(Mandatory = $false)]
		[string]
		$getUserType,
        [parameter(Mandatory = $false)]
        [string] 
        $System_Access_Flag
        )
       
                if($getUserType -eq "MailContact" -or $getUserType -eq "MADS")
                {
							#Check If tranisitondate is set in RF for user transitions
                            $msExchTransitionDate=(Get-ADUser $UserObjectID -Properties extensionAttribute10 -Server $domainController -Credential $Credential).extensionAttribute10
                            if([string]::IsNullOrEmpty($msExchTransitionDate))
                            {
                                    if($System_Access_Flag -eq "Y")
                                    {
                                        $userHashTable.Add("msExchRecipientDisplayType", $msExchRecipientDisplayType_Contact) 
                                        $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithSID) 
                                    }
                                    else
                                    {
                                        $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithoutSID) 
                                    }
                             }
                            
                }
                elseif($getUserType -eq "RemoteMailbox")
                {
                                $userHashTable.Add("msExchRecipientDisplayType", $msExchRecipientDisplayType_RemoteMailbox) 
                                $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_RemoteMailbox) 
                         }
                elseif($getUserType -eq "LocalMailbox")
                {
                                $userHashTable.Add("msExchRecipientDisplayType", $msExchRecipientDisplayType_LocalMailbox) 
                                $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_LocalMailbox) 
                         }

                                           
                  
                    
}

function addItemToHashTable($hashTable, $key, $value)
{

if(($value -ne $null) -and (!$hashTable.Contains($key))){$hashTable.Add($key, $value) }
}


#This method sets the required user and mail attribute for the users based on the updated/deleted attributes
#HCM Comments - Need to check if these can be implemented in RF MA
function setADObject_ConnectorAttributes
{
	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)
	"setADObject_ConnectorAttributes:  Updating all modified or deleted rf attributes in userHashTable and mailBoxPropertiesHashTable">>$logFileLocation
    $modifiedAttributes=$CSEntryChange.AttributeChanges.Name
	#Identify all modified attributes in CS
    foreach ($modifiedAttribute in $modifiedAttributes)
    {
            $userValues= Get-ADUser $UserObjectID -Properties * -Server $domainController -Credential $Credential
            $msdsPhoneticDisplayName=$userValues.'msDS-PhoneticDisplayName'
            $msExchRTD=$userValues.msExchRecipientTypeDetails
            $fName=$null
            $lName=$null
            if($msdsPhoneticDisplayName -ne $null)
            {
                $fName=$msdsPhoneticDisplayName.split(' ')[0]
                $lName=$msdsPhoneticDisplayName.split(' ')[1]
            }

        if($CSEntryChange.AttributeChanges[$modifiedAttribute].ModificationType -ne "Delete")
        {            
			"setADObject_ConnectorAttributes: Checking all updated or added attributes in CS">>$logFileLocation
            switch($modifiedAttribute)
            {
                "deprovisionedDate"
                {
					#check for Mailcontact types
					#HCM comments - This logic may change based on mailcontact implementation
                    if($msExchRTD -eq $msExchRecipientTypeDetails_ContactWithSID -or $msExchRTD -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                    {
					#If deporvision date is present for mailcontact set targetaddress to inactiveforwarder.com
					"setADObject_ConnectorAttributes: Modified Attr is Deprovision Date. Setting targetaddress to inactiveforwarder.com for mailcontact">>$logFileLocation
					
                        $dummytargetAddress= "SMTP:" + $UserObjectID + "@" + $bogusdomain
                        addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $dummytargetAddress
                        
                    }
                    
					#Hide the address in Exchange
                    $authOrig="CN=$UserObjectID,"+$parentOU
                    addItemToHashTable $userHashTable "msExchHideFromAddressLists" $true
                    addItemToHashTable $userHashTable "msExchRequireAuthToSendTo" $true
                    addItemToHashTable $userHashTable "authOrig" $authOrig
                    break
                }
                "anglczd_first_nm"
                {
                    $fName=Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "anglczd_first_nm"
                    break
                }
                "anglczd_last_nm"
                {
                    $lName=Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "anglczd_last_nm"
                    break
                }
                "getUserType"
                {
                    break
                }
                "employeeID"
                {
                    break
                }
              
                "mailNickname"
                {
                    $mailNickname= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "mailNickname"
                    addItemToHashTable $mailBoxPropertiesHashTable "mailNickname" $mailNickname
                    break
                }
                "domain"
                {
                break
            
                }
                "targetAddress"
                {
                    $targetAddress= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "targetAddress"
                    addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $targetAddress
                    break

                }
                "businessAreaCode"
                {
                    $businessAreaCode= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "businessAreaCode"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute7" $businessAreaCode
                    break
                }
                "employeeType"
                {
                    $employeeType= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeType"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute6" $employeeType
                    break
                }
                "org_unit_cd"
                {
                    $orgUnitCode= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "org_unit_cd"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute3" $orgUnitCode
                    break
                }
                "supervisor_flag"
                {
                    $supervisor_flag= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "supervisor_flag"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute1" $supervisor_flag
                    break
                }
                "employeeStatus"
                {
                    $employeeStatus= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeStatus"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute2" $employeeStatus
                    break
                }
                "personnel_area_cd"
                {
                    $personnel_area_cd= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "personnel_area_cd"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute4" $personnel_area_cd
                    break
                }
                "cost_center_cd"
                {
                    $cost_center_cd= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "cost_center_cd"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute5" $cost_center_cd
                    break
                }
                "company_code"
                {
                    $company_code= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "company_code"
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute13" $company_code
                    break
                }
				#Set UserAccountControl based on SA flag
                "System_Access_Flag"
                {
                    if($System_Access_Flag -eq $true)
                    {
                        addItemToHashTable $userHashTable "msExchUserAccountControl" 2

                    }
                    else
                    {
                        addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                    }
                    break
                }

            }        
        }
      
		#If the attributes are Deleted.
		#HCM comments -  these update can be moved to RF MA flows.
        else
        {
			"setADObject_ConnectorAttributes: Checking all Deleted attributes in CS">>$logFileLocation
            switch($modifiedAttribute)
            {
                "deprovisionedDate"
                {

                    if($msExchRTD -eq $msExchRecipientTypeDetails_ContactWithSID -or $msExchRTD -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                    {
					"setADObject_ConnectorAttributes: Deleted Attr is Deprovision Date. Restoring targetaddress from updateRecipientFlag">>$logFileLocation
						#HCM Comments -  Verify the UpdateRecepientFlag value details
                        $valuesFromSAD= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "updateRecipientFlag"
                        if($valuesFromSAD -ne $null)
                        {
                            $RestoredTargetAddress=$valuesFromSAD.Split(',')[4].subString("RestoredTargetAddress=".Length)
                        }
                        $RestoredTargetAddress="SMTP:" + $RestoredTargetAddress
                        addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $RestoredTargetAddress
                    }
                    $userObject.msExchHideFromAddressLists=$null
                    $userObject.msExchRequireAuthToSendTo=$null
                    $userObject.authOrig=$null

                    break
                    }
                "anglczd_last_nm"
                {
                    $lName=$null
                    break
                }
                "msExchHomeServerName"
                {
                    break
                }
                "homeMDB"
                {
                    break
                }
                "mailNickname"
                {
                    break
                }
                "domain"
                {
                break
            
                }
                "targetAddress"
                {

                    break

                }
                "businessAreaCode"
                {
                    $userObject.extensionAttribute7=$null
                    break
                }
                "employeeType"
                {
                    $userObject.extensionAttribute6=$null
                    break
                }
                "org_unit_cd"
                {
                    $userObject.extensionAttribute3=$null
                    break
                }
                "supervisor_flag"
                {
                    $userObject.extensionAttribute1=$null
                    break
                }
                "employeeStatus"
                {
                    $userObject.extensionAttribute2=$null
                    break
                }
                "personnel_area_cd"
                {
                    $userObject.extensionAttribute4=$null
                    break
                }
                "cost_center_cd"
                {
                    $userObject.extensionAttribute5=$null
                    break
                }
                "company_code"
                {
                    $userObject.extensionAttribute13=$null
                    break
                }
                "System_Access_Flag"
                {
                    addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                    break
                }

            }
        }
        $phoneticName=$fName
        if($lName -ne $null){$phoneticName=$fName + " " + $lName}
        
        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $phoneticName
        

   
    }
}

 ##region Exch_Trans_CD
	#Code to determine exchange code 
	#Exch_Trans_CD values 
	#0 - None (Do nothing)
	#1 - Create Mailbox (after tipping point) 
	#2 - Maintain Mailbox
	#3 - Unhide Mailbox
	#4 - Create Contact (after tipping point) - SAD Auth and Ext Contact
	#5 - Maintain Contact 
	#6 - Clear Mailbox properties
	#7 - Mailbox to Contact (day 0)
	#8 - Maintain Contact - after transition from mailbox (within 7 days)
	#9 - Clear all mail properties
	#10 - Mailbox to none (day 0) - None
function setADObject_ExchTransCode
{

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)
		"setADObject_ExchTransCode: Updating mailbox, mailcontact attributes based on ExchTransCode">>$logFileLocation
        $exchangeTransitionCode= Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "exchangeTransitionCode" 
        $getUserType= Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "getUserType" 
        $existingUserDetails=Get-ADUser $UserObjectID -Properties employeeID,msExchRecipientTypeDetails,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15,mailNickname,userPrincipalName,msExchRecipientTypeDetails,msExchMasterAccountSid,msExchUserAccountControl,msRTCSIP-PrimaryUserAddress,mail,targetAddress -Server $domainController -Credential $Credential


		#HCM Comments -  Update mailbox/Mailcontact properties based on Exchange Transition codes set in SAD MA.
        switch ($exchangeTransitionCode)
        {
        "1"
           {	
				"setADObject_ExchTransCode: Creating mailbox for exchangeTransitionCode=1">>$logFileLocation
				#Create mailbox
				#HCM Comments- Mailcontact to Mailbox User transition
				#Disable mailcontact user . Need to update the logic for mailcontact objects
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithSID -or $existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                {
					"setADObject_ExchTransCode: Updating mailBoxPropertiesHashTable from RF for mailcontact to mailbox transition at exchangeTransitionCode=1">>$logFileLocation
					
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute1" $existingUserDetails.extensionAttribute1
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute2" $existingUserDetails.extensionAttribute2
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute3" $existingUserDetails.extensionAttribute3
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute4" $existingUserDetails.extensionAttribute4
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute5" $existingUserDetails.extensionAttribute5
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute6" $existingUserDetails.extensionAttribute6
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute7" $existingUserDetails.extensionAttribute7
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute8" $existingUserDetails.extensionAttribute8
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute9" $existingUserDetails.extensionAttribute9
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute10" $existingUserDetails.extensionAttribute10
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute11" $existingUserDetails.extensionAttribute11
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute12" $existingUserDetails.extensionAttribute12
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute13" $existingUserDetails.extensionAttribute13
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute14" $existingUserDetails.extensionAttribute14
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute15" $existingUserDetails.extensionAttribute15
                    addItemToHashTable $mailBoxPropertiesHashTable "mailNickname" $existingUserDetails.mailNickname
                    
                    $null=disable-mailUser $UserObjectID -DomainController $domainController -Confirm:$false
                }
					"setADObject_ExchTransCode: Updating LegacyDN in userHashTable from RF for mailcontact to mailbox transition at exchangeTransitionCode=1">>$logFileLocation
                    $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                    addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN

                switch($getUserType)
                {
                    "LocalMailbox"
                    {
					"setADObject_ExchTransCode: LocalMailbox at exchangeTransitionCode=1,Setting HomeMDB value ">>$logFileLocation
                    #region Random selection of homeMDB/msExchHomeServerName
                        $ExchHomeServer = $parentnodelist | where {$_.PERSONNELAREA.CODE -eq  $existingUserDetails.extensionAttribute4 -and $_.CONSTITUENT.EMPLY_GRP -eq  $existingUserDetails.extensionAttribute6}
                        if($ExchHomeServer.REGION.ExchHomeServer.Attributes.Count -gt 1)
                        {
                            $randomNum_ExchServers=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer.Attributes.Count
                            $ExchHomeServerName_RR=$ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].name
                            $randomNum_homeMDB=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].homeMDB.Count 
                            $homeMDB_RR=$ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].homeMDB[$randomNum_homeMDB]
                        }
                        else
                        {
                            $ExchHomeServerName_RR=$ExchHomeServer.REGION.ExchHomeServer.name
                            $randomNum_homeMDB=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer.homeMDB.Count
                            $homeMDB_RR=$ExchHomeServer.REGION.ExchHomeServer.homeMDB[$randomNum_homeMDB].ToString()
                        }

                    #endregion 

						"setADObject_ExchTransCode: LocalMailbox at exchangeTransitionCode=1, Updating UserHasTable and mailBoxPropertiesHashTable ">>$logFileLocation
						addItemToHashTable $userHashTable "msExchRBACPolicyLink" $RBAC
						addItemToHashTable $userHashTable "mDBuseDefaults" $true
						addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $existingUserDetails.'msDS-PhoneticDisplayName'


						addItemToHashTable $mailBoxPropertiesHashTable "homeMDB" $homeMDB_RR
						addItemToHashTable $mailBoxPropertiesHashTable "msExchHomeServerName" $ExchHomeServerName_RR

                    break
                    }
                    "RemoteMailbox"
                    {
						"setADObject_ExchTransCode: RemoteMailbox at exchangeTransitionCode=1, Updating UserHasTable ">>$logFileLocation
						
                        addItemToHashTable $userHashTable "msExchRBACPolicyLink" $RBAC
                        addItemToHashTable $userHashTable "mDBuseDefaults" $true
                        addItemToHashTable $userHashTable "msExchUserAccountControl" 2
                        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $existingUserDetails.'msDS-PhoneticDisplayName'


                    break
                    }
                }
                break
           }

        "2"
            {
				#Mailtain Mailbox
				"setADObject_ExchTransCode: Maintain mailbox at exchangeTransitionCode=2">>$logFileLocation
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailboxwithoutSID -or  $existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
					"setADObject_ExchTransCode: Localmailbox at exchangeTransitionCode=2 , Remove targetaddress">>$logFileLocation
                    $mailBoxUserObject.targetAddress=$null
                }
                break
            }
        "3"
            {
				#Unhide mailbox for none/mailcontact to mailbox transition or reprovisioning scenarios
				"setADObject_ExchTransCode: exchangeTransitionCode=3, Unhide mailbox for none/mailcontact to mailbox transition or reprovisioning scenarios">>$logFileLocation
                $userObject.msExchHideFromAddressLists=$null
                $userObject.msExchRequireAuthToSendTo=$null
                $userObject.authOrig=$null

                break
            }
        "4"
            {
				# Create Contact - Adding required attributes to create mail contact.
				#HCM comments -  this logic will change for mailcontact objects
				"setADObject_ExchTransCode: exchangeTransitionCode=4, Create Contact - Adding required attributes to create mail contact.">>$logFileLocation
                addItemToHashTable $userHashTable "msExchRBACPolicyLink" $existingUserDetails.'msDS-PhoneticDisplayName'
                $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN

                break
            }
        "5"
            {
			#Mailntain contact
				"setADObject_ExchTransCode: exchangeTransitionCode=5, Maintain Contact - Do Nothing.">>$logFileLocation
                break
            }
        "6"
            {
				"setADObject_ExchTransCode: exchangeTransitionCode=6, Clear mailbox properties.">>$logFileLocation
				#Clear mailbox properties.
				#HCM Comments- Below EA attributes can be set in RF MA.
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute1" $existingUserDetails.extensionAttribute1
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute2" $existingUserDetails.extensionAttribute2
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute3" $existingUserDetails.extensionAttribute3
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute4" $existingUserDetails.extensionAttribute4
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute5" $existingUserDetails.extensionAttribute5
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute6" $existingUserDetails.extensionAttribute6
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute7" $existingUserDetails.extensionAttribute7
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute8" $existingUserDetails.extensionAttribute8
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute9" $existingUserDetails.extensionAttribute9
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute10" $existingUserDetails.extensionAttribute10
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute11" $existingUserDetails.extensionAttribute11
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute12" $existingUserDetails.extensionAttribute12
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute13" $existingUserDetails.extensionAttribute13
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute14" $existingUserDetails.extensionAttribute14
                    addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute15" $existingUserDetails.extensionAttribute15
                   
				   "setADObject_ExchTransCode: exchangeTransitionCode=6, Setting mailboxpropertieshashtable with RF attributes required to disable mailbox">>$logFileLocation
				   
					addItemToHashTable $mailBoxPropertiesHashTable "mailNickname" $existingUserDetails.mailNickname
					addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $existingUserDetails.targetAddress
                
				"setADObject_ExchTransCode: exchangeTransitionCode=6, Calling Set-MsolUserLicense to Remove Licenses of mailbox users ">>$logFileLocation
				#HCM Comments - For HCM below logic to remove license can be removed
                $licenses=Get-MsolUser -UserPrincipalName $existingUserDetails.userPrincipalName -ErrorAction Ignore| Select-Object licenses
                if($licenses -ne $null)
                {
                    if($licenses.Licenses.Count -ne 0)
                    {
                        Set-MsolUserLicense -UserPrincipalName $existingUserDetails.userPrincipalName -RemoveLicenses $licenses.licenses.AccountSkuID -ErrorAction Ignore
                    }
                }
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
				"setADObject_ExchTransCode: exchangeTransitionCode=6, Calling disable Remotemailbox for RemoteMailbox user ">>$logFileLocation
                    $null=Disable-RemoteMailbox $UserObjectID -DomainController $domainController -confirm:$False
                        
                }
                elseif($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
				"setADObject_ExchTransCode: exchangeTransitionCode=6, Calling disable Remotemailbox for LocalMailbox user ">>$logFileLocation
                        $null=Disable-Mailbox $UserObjectID -DomainController $domainController -confirm:$False
                }
              
				# Add try catch for above disable commands.
				
                addItemToHashTable $userHashTable "msExchUserAccountControl" $existingUserDetails.msExchUserAccountControl
                addItemToHashTable $userHashTable "msRTCSIP-PrimaryUserAddress" $existingUserDetails.'msRTCSIP-PrimaryUserAddress'
                $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN

				#HCM Comments - Need to check if recpientType details values can be set in RF or we need addtional error check before we set these values
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox -or  $existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
					"setADObject_ExchTransCode: exchangeTransitionCode=6, update recpientTypeDetails in userHashTable after disable mailbox for Remote and LocalMailbox">>$logFileLocation
                    $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithSID) 
                    $userHashTable.Add("msExchRecipientDisplayType", $msExchRecipientDisplayType_Contact) 
                }
                elseif($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailboxwithoutSID)
                {
					"setADObject_ExchTransCode: exchangeTransitionCode=6, update recpientTypeDetails for mailcontact">>$logFileLocation
                    $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithoutSID)
                }
                break
            }
        "7"
            {
			#Mailbox to Contact (day 0)
			"setADObject_ExchTransCode: exchangeTransitionCode=7, Mailbox to Contact (day 0), Do Nothing">>$logFileLocation
        break
            }
        "8"
            {
				#Maintain Contact - after transition from mailbox (within 7 days)
				"setADObject_ExchTransCode: exchangeTransitionCode=8, Maintain Contact - after transition from mailbox (within 7 days), Do Nothing">>$logFileLocation
                break
            }
        "9"
            {
			#Clear all mail properties
				"setADObject_ExchTransCode: exchangeTransitionCode=9, Clear all mail properties">>$logFileLocation
                $mailNickName = $existingUserDetails.mail.Split('@')[0]
                addItemToHashTable $mailBoxPropertiesHashTable "mailNickname" $mailNickName

				#Remove Licenses
                $licenses=Get-MsolUser -UserPrincipalName $existingUserDetails.userPrincipalName -ErrorAction Ignore| Select-Object licenses
                if($licenses -ne $null)
                {
                    if($licenses.Licenses.Count -ne 0)
                    {
						"setADObject_ExchTransCode: exchangeTransitionCode=9, Remove all licenses">>$logFileLocation
                        Set-MsolUserLicense -UserPrincipalName $existingUserDetails.userPrincipalName -RemoveLicenses $licenses.licenses.AccountSkuID -ErrorAction Ignore
                    }
                }
					"setADObject_ExchTransCode: exchangeTransitionCode=9, uPDTE UserHasTable WITH rf Attributes">>$logFileLocation
					
                    addItemToHashTable $userHashTable "extensionAttribute1" $existingUserDetails.extensionAttribute1
                    addItemToHashTable $userHashTable "extensionAttribute2" $existingUserDetails.extensionAttribute2
                    addItemToHashTable $userHashTable "extensionAttribute3" $existingUserDetails.extensionAttribute3
                    addItemToHashTable $userHashTable "extensionAttribute4" $existingUserDetails.extensionAttribute4
                    addItemToHashTable $userHashTable "extensionAttribute5" $existingUserDetails.extensionAttribute5
                    addItemToHashTable $userHashTable "extensionAttribute6" $existingUserDetails.extensionAttribute6
                    addItemToHashTable $userHashTable "extensionAttribute7" $existingUserDetails.extensionAttribute7
                    addItemToHashTable $userHashTable "extensionAttribute8" $existingUserDetails.extensionAttribute8
                    addItemToHashTable $userHashTable "extensionAttribute9" $existingUserDetails.extensionAttribute9
                    addItemToHashTable $userHashTable "extensionAttribute10" $existingUserDetails.extensionAttribute10
                    addItemToHashTable $userHashTable "extensionAttribute11" $existingUserDetails.extensionAttribute11
                    addItemToHashTable $userHashTable "extensionAttribute12" $existingUserDetails.extensionAttribute12
                    addItemToHashTable $userHashTable "extensionAttribute13" $existingUserDetails.extensionAttribute13
                    addItemToHashTable $userHashTable "extensionAttribute14" $existingUserDetails.extensionAttribute14
                    addItemToHashTable $userHashTable "extensionAttribute15" $existingUserDetails.extensionAttribute15
					
					"setADObject_ExchTransCode: exchangeTransitionCode=9, Set homeMDB and msExchHomeServerName NUll for the user">>$logFileLocation
                    $userObject.homeMDB=$null
                    $userObject.msExchHomeServerName=$null

               

        break
            }
        "10"
            {
				"setADObject_ExchTransCode: exchangeTransitionCode=10, mailbox to None day 0">>$logFileLocation
				# mailbox to None day 0
                $authOrig="CN=$UserObjectID,"+$parentOU
                addItemToHashTable $userHashTable "msExchHideFromAddressLists" $true
                addItemToHashTable $userHashTable "msExchRequireAuthToSendTo" $true
                addItemToHashTable $userHashTable "authOrig" $authOrig
				
				"setADObject_ExchTransCode: exchangeTransitionCode=10, hiding mailbox from Address list">>$logFileLocation

        break
            }

        "0"
        {
		#None type user
			"setADObject_ExchTransCode: exchangeTransitionCode=0, None Type user, Setting mailnickname to null">>$logFileLocation
            $userObject.mailNickname=$null
            break
        }
        }
}

#This mehtod sets the RF attributes requried for mailbox enable and disable in UserHasTable and mailBoxPropertiesHashTable 
function setADObject_UserType
{

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[string]

		$getUserType

	)

                switch ($getUserType)
                {

                    "LocalMailbox"
                    {
                    #region Random selection of homeMDB/msExchHomeServerName
                        $ExchHomeServer = $parentnodelist | where {$_.PERSONNELAREA.CODE -eq  $personnel_area_cd -and $_.CONSTITUENT.EMPLY_GRP -eq  $employeeType}
                        if($ExchHomeServer.REGION.ExchHomeServer.Attributes.Count -gt 1)
                        {
                            $randomNum_ExchServers=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer.Attributes.Count
                            $ExchHomeServerName_RR=$ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].name
                            $randomNum_homeMDB=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].homeMDB.Count 
                            $homeMDB_RR=$ExchHomeServer.REGION.ExchHomeServer[$randomNum_ExchServers].homeMDB[$randomNum_homeMDB]
                        }
                        else
                        {
                            $ExchHomeServerName_RR=$ExchHomeServer.REGION.ExchHomeServer.name
                            $randomNum_homeMDB=Get-Random -Minimum 0 -Maximum $ExchHomeServer.REGION.ExchHomeServer.homeMDB.Count
                            $homeMDB_RR=$ExchHomeServer.REGION.ExchHomeServer.homeMDB[$randomNum_homeMDB].ToString()
                        }

                    #endregion 

						"setADObject_UserType:  setting mailboxpropertieshashtable for local mailbox user">>$logFileLocation
                        addItemToHashTable $mailBoxPropertiesHashTable "msExchHomeServerName" $ExchHomeServerName_RR
                        addItemToHashTable $mailBoxPropertiesHashTable "homeMDB" $homeMDB_RR

						"setADObject_UserType:  setting userHashTable for local mailbox user">>$logFileLocation
                        addItemToHashTable $userHashTable "legacyExchangeDN" $legacyExchangeDN
                        addItemToHashTable $userHashTable "msExchRBACPolicyLink" $RBAC
                        addItemToHashTable $userHashTable "mDBuseDefaults" $true

                        $phoneticDisplayName=$anglczd_first_nm +" " + $anglczd_last_nm
                        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $phoneticDisplayName                      

                        
                       
                        if($System_Access_Flag -eq "Y")
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 2
                        }
                        else
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                        }
                       
                    break
                    }
                    "RemoteMailbox"
                    {

						"setADObject_UserType:  setting userHashTable for Remote mailbox user">>$logFileLocation
                        addItemToHashTable $userHashTable "legacyExchangeDN" $legacyExchangeDN
                        addItemToHashTable $userHashTable "msExchRBACPolicyLink" $RBAC
                        addItemToHashTable $userHashTable "mDBuseDefaults" $true

                        
                        $phoneticDisplayName=$anglczd_first_nm +" " + $anglczd_last_nm
                        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $phoneticDisplayName

                        
                        if($System_Access_Flag -eq "Y")
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 2

                        }
                        else
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                        }

                        
                    break
                    }
                    "MADS"
                    {
					
						"setADObject_UserType:  setting mailboxpropertieshashtable for MADS">>$logFileLocation
                        addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $targetAddress
                        
                        addItemToHashTable $userHashTable "legacyExchangeDN" $legacyExchangeDN
                        addItemToHashTable $userHashTable "mailNickname" $mailNickname
                        
                        $phoneticDisplayName=$anglczd_first_nm +" " + $anglczd_last_nm
                        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $phoneticDisplayName
                        
                        if($System_Access_Flag -eq "Y")
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 2

                        }
                        else
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                        }
                        
                    break
                     
                    }
                    "MailContact"
                    {
						"setADObject_UserType:  setting mailboxpropertieshashtable for mailcontact user">>$logFileLocation
                        addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $targetAddress

						"setADObject_UserType:  setting userHashTable for mailcontact user">>$logFileLocation
                        addItemToHashTable $userHashTable "legacyExchangeDN" $legacyExchangeDN
                        addItemToHashTable $userHashTable "mailNickname" $mailNickname
                        
                        $phoneticDisplayName=$anglczd_first_nm +" " + $anglczd_last_nm
                        addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $phoneticDisplayName
                        
                        
                        if($System_Access_Flag -eq "Y")
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 2

                        }
                        else
                        {
                            addItemToHashTable $userHashTable "msExchUserAccountControl" 0
                        }
                        
                    break
                    }
                    "None"
                    {
                      break
                    }

                    default
                    {
                    break
                    }
                }

}


#This is fucntion startup function of Exchange Export MA. It modifies the User object attributes based on the CSEntryChanges
function Export-CSEntries

{

	<#

	.Synopsis

		Exports the CSEntry changes.

	.Description

		Exports the CSEntry changes.

	#>

	

	[CmdletBinding()]

	[OutputType([System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChangeResult]])]

	param(

	)



	$csentryChangeResults = New-GenericObject System.Collections.Generic.List Microsoft.MetadirectoryServices.CSEntryChangeResult

    (Get-Date).ToShortDateString() + "Evaluating each CSEntry Change. Method - Export-CSEntries">>$logFileLocation
	#loop through all CSEntries from CS and exceute operations for mailbox/mailcontact creation or tranisiton based on each CS entry attributes validations
	foreach ($csentryChange in $CSEntries)

	{        

		$Error.Clear() = $null

		$newAnchorTable = @{}
		
		#read all the attributes from CS required for user mailbox/mailcontact validations
		
		$dn = Get-CSEntryChangeDN $csentryChange
        $UserObjectID = $csentryChange.DN.ToString().Substring('USER='.Length)
		
		(Get-Date).ToShortDateString() + ":- $UserObjectID : Method - Export-CSEntries">>$logFileLocation
		
        $employeeID = Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeID" 
		
		#read default legacyExchangeDN from config
        $legacyExchangeDN=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/Exchange/legacyExchangeDN").InnerText + $employeeID
		
        "Method - Export-CSEntries : Reading attributes from CS">>$logFileLocation
		
        $msExchUserAccountControl=2
        $mDBuseDefaults=$true
        $msExchHomeServerName=Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "msExchHomeServerName"
        $homeMDB=Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "homeMDB"
        $anglczd_first_nm= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "anglczd_first_nm" 
        $anglczd_last_nm= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "anglczd_last_nm"
        $mailNickname= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "mailNickname" 
        $targetAddress= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "targetAddress" 
        $exchangeTransitionCode = Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "exchangeTransitionCode"
        $orgUnitCode = Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "org_unit_cd"
        $whereToProvision= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "whereToProvision"
        $deprovisionedDate= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "deprovisionedDate"
        $businessAreaCode= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "businessAreaCode"
        $employeeType= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeType"
        $updateRecipientFlag= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "updateRecipientFlag"
        $supervisor_flag= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "supervisor_flag"
        $employeeStatus= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeStatus"
        $personnel_area_cd= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "personnel_area_cd"
        $cost_center_cd= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "cost_center_cd"
        $company_code= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "company_code"
        $userType= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "getUserType"
        $SAFlag= Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "System_Access_Flag"
        
        $getUserType=$null
        $System_Access_Flag=$null
        $RestoredTargetAddress=$null
		
        if($updateRecipientFlag -ne $null)# Read required attributes from UpdateReceoientFlag set in SAD MA
        {
        $getUserType=$updateRecipientFlag.Split(',')[1].subString("userType=".Length)
        $System_Access_Flag=$updateRecipientFlag.Split(',')[2].subString("systemAccessFlag=".Length)
        $RestoredTargetAddress=$updateRecipientFlag.Split(',')[4].subString("RestoredTargetAddress=".Length)
        }

     
  
		$objectType = $csentryChange.ObjectType

		$objectModificationType = $csentryChange.ObjectModificationType

		Write-Debug "Exporting $objectModificationType to $objectType : $dn"
        
		"Method - Export-CSEntries" + "Exporting $objectModificationType to $objectType : $dn">>$logFileLocation		


		try

		{

			switch ($objectType)

			{

				"User"

				{
					#Call Export-User method to create new export anchor table for each CS entry changes
					(Get-Date).ToShortDateString() + ":- $UserObjectID : Method - Export-CSEntries . " + "Calling method Export-User">>$logFileLocation
                    $newAnchorTable = Export-User $csentryChange

					break

				}

				default

				{

					throw "Unknown CSEntry ObjectType: $_"

				}

			}

		}

		catch

		{

			Write-Error "$_"

		}

        

		if ($Error)

		{

			$csentryChangeResult = New-CSEntryChangeExportError -CSEntryChangeIdentifier $csentryChange.Identifier -ErrorObject $Error
            "$UserObjectID : "+(Get-Date).ToShortDateString()+":" +(Get-Date).ToShortTimeString() +": " + $Error>>$logFileLocation

		}

		else

		{

			$exportAdd = $objectModificationType -eq "Add"

            (Get-Date).ToShortDateString() + ":- $UserObjectID : Method - Export-CSEntries. " + "Adding new Anchor table for Export to change result set">>$logFileLocation
            #Adding new Anchor table for Export to change result set"
			$csentryChangeResult = New-CSEntryChangeResult -CSEntryChangeIdentifier $csentryChange.Identifier -NewAnchorTable $newAnchorTable -ExportAdd:$exportAdd



			Write-Debug "Exported $objectModificationType to $objectType : $dn."

		}



		$csentryChangeResults.Add($csentryChangeResult)

	}


	$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"

	return [Activator]::CreateInstance($closedType, $csentryChangeResults)

}

#This function evaluates the user attributes and mailbox properties updates for each CSEntryChange Modification type (i.e. Add, update, delete)
function Export-User

{

	<#

	.Synopsis

		Exports the changes for User objects.

	.Description

		Exports the changes for User objects.

		Returns the Hashtable of anchor attribute values for Export-Add csentries.

	#>

	

	[CmdletBinding()]

	[OutputType([Hashtable])]

	param(

		[parameter(Mandatory = $true)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)



	$newAnchorTable = @{}

	(Get-Date).ToShortDateString() + ":- $UserObjectID : Method - Export-User: Evaluating attributes for each modification type ">>$logFileLocation

	switch ($CSEntryChange.ObjectModificationType)

	{

		"Add"

		{
			"Export-User: Evaluating ModificationType ADD ">>$logFileLocation
            $Error.Clear() = $null
			$dn  = Get-CSEntryChangeDN $CSEntryChange

            $mailBoxPropertiesHashTable = New-Object HashTable
            $userHashTable = New-Object HashTable
            addItemToHashTable $userHashTable "sAMAccountName" $UserObjectID
			
			"Export-User - ADD: Adding  EA attributes to hashtable">>$logFileLocation
			#HCM comments - below EA attributes can flow from MV to RF directly
            addItemToHashTable $mailBoxPropertiesHashTable "sAMAccountName" $UserObjectID
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute1" $supervisor_flag
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute2" $employeeStatus
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute3" $orgUnitCode
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute4" $personnel_area_cd
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute5" $cost_center_cd
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute6" $employeeType
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute7" $businessAreaCode
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute13" $company_code
            
			"Export-User -ADD:  Calling Invoke-NewUserCommand method">>$logFileLocation
			# Call Invoke-NewUserCommand method to set the required AD attributes set in mailBoxPropertiesHashTableand userHashTable for a mailbox/mailcontact user (in setADObject_UserType method)
			Invoke-NewUserCommand $getUserType 
			
			"Export-User -ADD:  setting mailboxpropertieshashtable in RF">>$logFileLocation
			"Export-User -ADD: $UserObjectID - mailBoxPropertiesHashTable : $mailBoxPropertiesHashTable">>$logFileLocation			
			#HCM - Need to add try catch .Below call sets mailBoxPropertiesHashTable with EA attributes set above and by Invoke-NewUserCommand for mailbox/mailcontact user.
            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
			
			#HCM - TBD. Need to create separate logic for mailcontact creation.
            if($getUserType -eq "LocalMailbox" -or $getUserType -eq "MailContact" -or $getUserType -eq "MADS")
            {
				"Export-User -ADD:  calling update-recepient for Localmailbox, mailcontact and MADS users">>$logFileLocation
                $null=Update-Recipient $UserObjectID -DomainController $domainController
            }
            elseif($getUserType -eq "RemoteMailbox")
            {
				"Export-User -ADD:  calling Enable-RemoteMailbox for Remotemailbox users">>$logFileLocation
                $RemoteRoutingAddress=$mailNickname + "@" + $UPNDomain

                $cmd = "Enable-RemoteMailbox '$UserObjectID' -RemoteRoutingAddress '$RemoteRoutingAddress' -domainController '$domainController'"
                Invoke-Expression -Command $cmd | Out-Null 
				
				#Invoke-Expression -Command "$cmd -ErrorAction silentlycontinue" | Out-Null	

            }
			
			#Try {
			#$cmd = "Enable-RemoteMailbox '$UserObjectID' -RemoteRoutingAddress '$RemoteRoutingAddress' -domainController '$domainController'"
                #Invoke-Expression -Command "$cmd -ErrorAction silentlycontinue" | Out-Null	
			#}
			#Catch {
			#"Mailbox server doesn't exist"
			#}

			"Export-User-ADD:  calling set_RecTypeDetails_RecDisplayType">>$logFileLocation
			#HCM - TBD .Need to identify the details for Mailcontact. Calling this function to set the required recepient type details for mailbox/mailcontact users
            set_RecTypeDetails_RecDisplayType -getUserType $getUserType -System_Access_Flag $System_Access_Flag
			
			"Export-User-ADD:  setting userHashTable">>$logFileLocation
			"Export-User-ADD: $UserObjectID - userHashTable : $userHashTable">>$logFileLocation
			#set required AD attributes for mailbox/mailcontact user.
            $null=Set-ADUser $UserObjectID -Replace $userHashTable -Credential $Credential -Server $domainController
			
			#try{
			   # Set-ADUser $User.samaccountname <#stuff#> -ErrorAction silentlycontinue
			#}
			#catch{
			 #   write-output"Error setting value for '$($User.samaccountname)'" |
			  #  out-file "c:\error.log" -Append
			#}
            
			break


		}
		"Update"

		{
		#HCM Comments -  Need to check if we should add more atributes for update validations
    if((isModified "exchangeTransitionCode") -or (isModified "deprovisionedDate")-or (isModified "anglczd_first_nm")-or (isModified "anglczd_last_nm") -or (isModified "employeeID") -or (isModified "mailNickname") -or (isModified "targetAddress") -or (isModified "businessAreaCode") -or (isModified "employeeType") -or (isModified "org_unit_cd") -or (isModified "System_Access_Flag") -or (isModified "supervisor_flag") -or (isModified "employeeStatus") -or (isModified "personnel_area_cd") -or (isModified "cost_center_cd") -or (isModified "company_code") -or (isModified "getUserType"))
    {

            $userObject = Get-ADUser $UserObjectID -Properties * -Server $domainController -Credential $Credential
            $mailBoxUserObject = Get-ADUser $UserObjectID -Properties * -Server $domainController -Credential $Credential
            $userHashTable = New-Object HashTable
            $mailBoxPropertiesHashTable = New-Object HashTable
            
			"Export-User-UPDATE: $UserObjectID Calling Invoke-UpdateUserCommand Method">>$logFileLocation
			#Call Invoke-UpdateUserCommand to Update the AD and Mail attributes of the user based on modified attributes and ExchTransCode
			# Sub functions evaluates the modified attributes and save the new values in hastable. Sub functions also process the mailbox/mailcontacts properties based on the the Exchange transition codes set in SAD MA.
			Invoke-UpdateUserCommand $CSEntryChange
            addItemToHashTable $userHashTable "sAMAccountName" $UserObjectID
            addItemToHashTable $mailBoxPropertiesHashTable "sAMAccountName" $UserObjectID

			"Export-User-UPDATE: $UserObjectID - mailBoxPropertiesHashTable :- $mailBoxPropertiesHashTable">>$logFileLocation
			"Export-User-UPDATE: $UserObjectID - Updating mailBoxPropertiesHashTable in AD before calling Enable/Disable mailbox or Update Recepient">>$logFileLocation
            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
            $null=Set-ADUser -Instance $mailBoxUserObject -Credential $Credential -Server $domainController
			
			#Remove mail properties for mailcontact to none transition
            if($exchangeTransitionCode -eq "9")
            {
				
                if($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
					"Export-User-UPDATE: $UserObjectID - Calling diasble mailbox to Clear mail properties for EA=9, mailbox to none transition">>$logFileLocation
                    $null=Disable-RemoteMailbox $UserObjectID -DomainController $domainController -confirm:$False
                        
                }
                elseif($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
						"Export-User-UPDATE: $UserObjectID - Calling diasble mailbox to Clear  mail properties for EA=9, localmailbox to none transition">>$logFileLocation
                        $null=Disable-Mailbox $UserObjectID -DomainController $domainController -confirm:$False
                }
                elseif($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithSID -or $userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                {
					   "Export-User-UPDATE: $UserObjectID - Calling diasble mailbox to Clear  mail properties for EA=9, mailcontact to none transition">>$logFileLocation
                        $null=Disable-MailUser $UserObjectID -DomainController $domainController -confirm:$False
                }
            }
            
			#Create new mailbox
            if($userType -eq "RemoteMailbox" -and $exchangeTransitionCode -eq "1")
            {
						"Export-User-UPDATE: $UserObjectID -  Enable mailbox for EA=1, Remotemailbox user">>$logFileLocation
                        $RemoteRoutingAddress=$userObject.mailNickname + "@" + $UPNDomain
                        $cmd = "Enable-RemoteMailbox '$UserObjectID' -RemoteRoutingAddress '$RemoteRoutingAddress' -domainController '$domainController'"
                        Invoke-Expression -Command $cmd | Out-Null 
                        
            }
            #elseif($userObject.mailNickname -ne $null -and $userObject.Description -eq $null -and $exchangeFlag -eq "TRUE")
			#HCM Comments - Identify the Scenario
            elseif($userObject.mailNickname -ne $null -and $exchangeFlag -eq "TRUE")
            {
						"Export-User-UPDATE: $UserObjectID - Calling Update Recipient for ExchangeFlag = true">>$logFileLocation
                        $null=Update-Recipient $UserObjectID -DomainController $domainController

            }
            
			"Export-User-UPDATE: $UserObjectID - Setting Update Recepient after enable/Disable mailbox">>$logFileLocation
            set_RecTypeDetails_RecDisplayType -getUserType $getUserType -System_Access_Flag $System_Access_Flag
            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
            
			 "Export-User-UPDATE: $UserObjectID -UserHashTable : - $userHashTable">>$logFileLocation
            "Export-User-UPDATE: $UserObjectID - Updating userHashTable in RF after enable/Disable mailbox">>$logFileLocation
            $null=Set-ADUser -Instance $userObject -Credential $Credential -Server $domainController
            $null=Set-ADUser $UserObjectID -Replace $userHashTable -Credential $Credential -Server $domainController
       } 


			break


		}





		default

		{

			throw "Unknown CSEntry ObjectModificationType: $_"

		}

	}



	return $newAnchorTable

}


function Invoke-NewUserCommand

{

	<#

	.Synopsis

		Invokes Enable-CsUser cmdlet on the specified user csentry.

	.Description

		Invokes Enable-CsUser cmdlet on the specified user csentry.

	#>

	

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $true)]

		[string]

		$getUserType

	)

        $newAnchorTable.Add("DN", $dn)    #Must Remain
		"Invoke-NewUserCommand:  setADObject_UserType to set EA attributes values in userHashTable">>$logFileLocation
        setADObject_UserType $getUserType

      

	}


function Invoke-ReplaceUserCommand

{

	<#

	.Synopsis

		Invokes Move-CsUser cmdlet on the specified user csentry.

	.Description

		Invokes Move-CsUser cmdlet on the specified user csentry.

	#>

	

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $true)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)

        #Put Code Here



}

# Sub functions evaluates the modified attributes and save the new values in hastable. Sub functions also evaluates the Excahnge transition codes.
function Invoke-UpdateUserCommand

{

	<#

	.Synopsis

		Invokes Move-CsUser cmdlet on the specified user csentry.

	.Description

		Invokes Move-CsUser cmdlet on the specified user csentry.

	#>

	

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)

    
              $modifiedAttribute=$CSEntryChange.AttributeChanges.Name

			  "Invoke-UpdateUserCommand:  Calling setADObject_ConnectorAttributes to set RF attributes values in userHashTable and mailBoxPropertiesHashTable">>$logFileLocation
              setADObject_ConnectorAttributes $CSEntryChange
			  
			  "Invoke-UpdateUserCommand:  Calling setADObject_ExchTransCode to set RF attributes values in userHashTable and mailBoxPropertiesHashTable">>$logFileLocation
              setADObject_ExchTransCode $CSEntryChange             


}


function Invoke-DisableUserCommand

{

	<#

	.Synopsis

		Invokes Disable-CsUser cmdlet on the specified user csentry.

	.Description

		Invokes Disable-CsUser cmdlet on the specified user csentry.

	#>

	

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $true)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)

        #Put Code Here



}

function Get-CsIdentity

{

	<#

	.Synopsis

		Gets the identifier for specified user csentry.

	.Description

		Gets the identifier for specified user csentry.

		It is the Guid if the Anchor is populated. Otherwise DN.

	#>

	

	[CmdletBinding()]

	[OutputType([string])]

	param(

		[parameter(Mandatory = $true)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)



	if (!$Error)

	{

		$dn  = Get-CSEntryChangeDN $CSEntryChange



		if ($CSEntryChange.AnchorAttributes.Contains("Guid") -and $CSEntryChange.AnchorAttributes["Guid"].Value -ne $null)

		{

			return ([Guid]$CSEntryChange.AnchorAttributes["Guid"].Value).ToString()

		}

		else # should only be here when ObjectModificationType = "Add"

		{

			return $dn.Replace("`'","''") # escape any single quotes in the DN

		}

	}

}

# Initial Method 
Export-CSEntries



Exit-Script -ScriptType "Export" -SuppressErrorCheck -ErrorObject $Error 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 

 
  
 
 
 
 
 
 
 
 

