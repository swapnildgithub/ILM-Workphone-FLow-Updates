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


#$CSEntryChange.AttributeChanges["deprovisionedDate"].ModificationType

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



function setADObject_ConnectorAttributes
{
	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)
    $modifiedAttributes=$CSEntryChange.AttributeChanges.Name
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
            

            switch($modifiedAttribute)
            {
                "deprovisionedDate"
                {
                    if($msExchRTD -eq $msExchRecipientTypeDetails_ContactWithSID -or $msExchRTD -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                    {
                        $dummytargetAddress= "SMTP:" + $UserObjectID + "@" + $bogusdomain
                        addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $dummytargetAddress
                        
                    }
                    <#
                    $licenses=Get-MsolUser -UserPrincipalName $userValues.userPrincipalName -ErrorAction Ignore| Select-Object licenses
                    if($licenses -ne $null)
                    {
                        if($licenses.Licenses.Count -ne 0)
                        {
                            Set-MsolUserLicense -UserPrincipalName $userValues.userPrincipalName -RemoveLicenses $licenses.licenses.AccountSkuID -ErrorAction Ignore
                        }
                    }
                    #>
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
      

        else
        {
            switch($modifiedAttribute)
            {
                "deprovisionedDate"
                {

                    if($msExchRTD -eq $msExchRecipientTypeDetails_ContactWithSID -or $msExchRTD -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                    {
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
function setADObject_ExchTransCode
{

	[CmdletBinding()]

	param(

		[parameter(Mandatory = $false)]

		[Microsoft.MetadirectoryServices.CSEntryChange]

		$CSEntryChange

	)
        $exchangeTransitionCode= Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "exchangeTransitionCode" 
        $getUserType= Get-CSEntryChangeValue -CSEntryChange $CSEntryChange -AttributeName "getUserType" 
        $existingUserDetails=Get-ADUser $UserObjectID -Properties employeeID,msExchRecipientTypeDetails,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15,mailNickname,userPrincipalName,msExchRecipientTypeDetails,msExchMasterAccountSid,msExchUserAccountControl,msRTCSIP-PrimaryUserAddress,mail,targetAddress -Server $domainController -Credential $Credential



        switch ($exchangeTransitionCode)
        {
        "1"
           {
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
                    
					 #"Set-ADUser $UserObjectID -clear msExchRecipientTypeDetails -Credential $Credential -Server $domainController">>$logFileLocation
					
					# Contact Release - Code update for Exchange 2016 - Clear msExchRecipientTypeDetails in RF before disabling mailcontact
					#$null=Set-ADUser $UserObjectID -clear msExchRecipientTypeDetails -Credential $Credential -Server $domainController						
					#$null=Set-ADUser -Instance $mailBoxUserObject -Credential $Credential -Server $domainController
					
					
					  "Disable-MailUser $UserObjectID -DomainController $domainController">>$logFileLocation
						
					
                    $null=disable-mailUser $UserObjectID -DomainController $domainController -Confirm:$false
                }
                    $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                    addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN

                switch($getUserType)
                {
                    "LocalMailbox"
                    {

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


                    addItemToHashTable $userHashTable "msExchRBACPolicyLink" $RBAC
                    addItemToHashTable $userHashTable "mDBuseDefaults" $true
                    addItemToHashTable $userHashTable "msDS-PhoneticDisplayName" $existingUserDetails.'msDS-PhoneticDisplayName'


                    addItemToHashTable $mailBoxPropertiesHashTable "homeMDB" $homeMDB_RR
                    addItemToHashTable $mailBoxPropertiesHashTable "msExchHomeServerName" $ExchHomeServerName_RR

                    break
                    }
                    "RemoteMailbox"
                    {

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
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailboxwithoutSID -or  $existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
                    $mailBoxUserObject.targetAddress=$null
                }
                break
            }
        "3"
            {
                $userObject.msExchHideFromAddressLists=$null
                $userObject.msExchRequireAuthToSendTo=$null
                $userObject.authOrig=$null

                break
            }
        "4"
            {
                addItemToHashTable $userHashTable "msExchRBACPolicyLink" $existingUserDetails.'msDS-PhoneticDisplayName'
                $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN

                break
            }
        "5"
            {
                break
            }
        "6"
            {
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
                addItemToHashTable $mailBoxPropertiesHashTable "targetAddress" $existingUserDetails.targetAddress
                
				#HCM Phase 1 - retirement. Commneting Licensing code.
				<#
                $licenses=Get-MsolUser -UserPrincipalName $existingUserDetails.userPrincipalName -ErrorAction Ignore| Select-Object licenses
                if($licenses -ne $null)
                {
                    if($licenses.Licenses.Count -ne 0)
                    {
                        Set-MsolUserLicense -UserPrincipalName $existingUserDetails.userPrincipalName -RemoveLicenses $licenses.licenses.AccountSkuID -ErrorAction Ignore
                    }
                }
				#>
                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
                    $null=Disable-RemoteMailbox $UserObjectID -DomainController $domainController -confirm:$False
                        
                }
                elseif($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
                        $null=Disable-Mailbox $UserObjectID -DomainController $domainController -confirm:$False
                }
              
                addItemToHashTable $userHashTable "msExchUserAccountControl" $existingUserDetails.msExchUserAccountControl
                addItemToHashTable $userHashTable "msRTCSIP-PrimaryUserAddress" $existingUserDetails.'msRTCSIP-PrimaryUserAddress'
                $legacyDN = $legacyExchangeDNWithoutEmpID + $existingUserDetails.EmployeeID
                addItemToHashTable $userHashTable "legacyExchangeDN" $legacyDN


                if($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox -or  $existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
                    $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithSID) 
                    $userHashTable.Add("msExchRecipientDisplayType", $msExchRecipientDisplayType_Contact) 
                }
                elseif($existingUserDetails.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailboxwithoutSID)
                {
                    $userHashTable.Add("msExchRecipientTypeDetails", $msExchRecipientTypeDetails_ContactWithoutSID)
                }
                break
            }
        "7"
            {

        break
            }
        "8"
            {
                break
            }
        "9"
            {
                $mailNickName = $existingUserDetails.mail.Split('@')[0]
                addItemToHashTable $mailBoxPropertiesHashTable "mailNickname" $mailNickName

				#HCM Phase 1 - retirement. Commenting Licensing code.
				<#
                $licenses=Get-MsolUser -UserPrincipalName $existingUserDetails.userPrincipalName -ErrorAction Ignore| Select-Object licenses
                if($licenses -ne $null)
                {
                    if($licenses.Licenses.Count -ne 0)
                    {
                        Set-MsolUserLicense -UserPrincipalName $existingUserDetails.userPrincipalName -RemoveLicenses $licenses.licenses.AccountSkuID -ErrorAction Ignore
                    }
                }
				#>
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
                    $userObject.homeMDB=$null
                    $userObject.msExchHomeServerName=$null

               

        break
            }
        "10"
            {

                $authOrig="CN=$UserObjectID,"+$parentOU
                addItemToHashTable $userHashTable "msExchHideFromAddressLists" $true
                addItemToHashTable $userHashTable "msExchRequireAuthToSendTo" $true
                addItemToHashTable $userHashTable "authOrig" $authOrig


        break
            }

        "0"
        {
            $userObject.mailNickname=$null
            break
        }
        }
}

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


                        addItemToHashTable $mailBoxPropertiesHashTable "msExchHomeServerName" $ExchHomeServerName_RR
                        addItemToHashTable $mailBoxPropertiesHashTable "homeMDB" $homeMDB_RR

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



	foreach ($csentryChange in $CSEntries)

	{
        

		$Error.Clear() = $null

		$newAnchorTable = @{}
		$dn = Get-CSEntryChangeDN $csentryChange
        $UserObjectID = $csentryChange.DN.ToString().Substring('USER='.Length)

        $employeeID = Get-CSEntryChangeValue -CSEntryChange $csentryChange -AttributeName "employeeID" 
        $legacyExchangeDN=$RulesConfig.SelectSingleNode("/rules-extension-properties/management-agents/$environment/Exchange/legacyExchangeDN").InnerText + $employeeID
        
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
        if($updateRecipientFlag -ne $null)
        {
        $getUserType=$updateRecipientFlag.Split(',')[1].subString("userType=".Length)
        $System_Access_Flag=$updateRecipientFlag.Split(',')[2].subString("systemAccessFlag=".Length)
        $RestoredTargetAddress=$updateRecipientFlag.Split(',')[4].subString("RestoredTargetAddress=".Length)
        }

     
  
		$objectType = $csentryChange.ObjectType

		$objectModificationType = $csentryChange.ObjectModificationType

		Write-Debug "Exporting $objectModificationType to $objectType : $dn"
        


		try

		{

			switch ($objectType)

			{

				"User"

				{
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

            
            
			$csentryChangeResult = New-CSEntryChangeResult -CSEntryChangeIdentifier $csentryChange.Identifier -NewAnchorTable $newAnchorTable -ExportAdd:$exportAdd



			Write-Debug "Exported $objectModificationType to $objectType : $dn."

		}



		$csentryChangeResults.Add($csentryChangeResult)

	}



	##$exportEntriesResults = New-Object -TypeName "Microsoft.MetadirectoryServices.PutExportEntriesResults" -ArgumentList $csentryChangeResults

	$closedType = [Type] "Microsoft.MetadirectoryServices.PutExportEntriesResults"

	return [Activator]::CreateInstance($closedType, $csentryChangeResults)

}


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

    

	switch ($CSEntryChange.ObjectModificationType)

	{

		"Add"

		{



            $Error.Clear() = $null
			$dn  = Get-CSEntryChangeDN $CSEntryChange

            $mailBoxPropertiesHashTable = New-Object HashTable
            $userHashTable = New-Object HashTable
            addItemToHashTable $userHashTable "sAMAccountName" $UserObjectID
            addItemToHashTable $mailBoxPropertiesHashTable "sAMAccountName" $UserObjectID
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute1" $supervisor_flag
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute2" $employeeStatus
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute3" $orgUnitCode
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute4" $personnel_area_cd
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute5" $cost_center_cd
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute6" $employeeType
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute7" $businessAreaCode
            addItemToHashTable $mailBoxPropertiesHashTable "extensionAttribute13" $company_code
            
			Invoke-NewUserCommand $getUserType 
            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
            if($getUserType -eq "LocalMailbox" -or $getUserType -eq "MailContact" -or $getUserType -eq "MADS")
            {
                $null=Update-Recipient $UserObjectID -DomainController $domainController
            }
            elseif($getUserType -eq "RemoteMailbox")
            {


                $RemoteRoutingAddress=$mailNickname + "@" + $UPNDomain

                $cmd = "Enable-RemoteMailbox '$UserObjectID' -RemoteRoutingAddress '$RemoteRoutingAddress' -domainController '$domainController'"
                Invoke-Expression -Command $cmd | Out-Null
                        


            }
            set_RecTypeDetails_RecDisplayType -getUserType $getUserType -System_Access_Flag $System_Access_Flag
            $null=Set-ADUser $UserObjectID -Replace $userHashTable -Credential $Credential -Server $domainController
            
			break


		}
		"Update"

		{
    #if($exchangeTransitionCode -ne $null -or ($deprovisionedDate -ne $null -or $CSEntryChange.AttributeChanges["deprovisionedDate"].ModificationType -eq "Delete") -or $anglczd_first_nm -ne $null -or ($anglczd_last_nm -ne $null -or $CSEntryChange.AttributeChanges["anglczd_last_nm"].ModificationType -eq "Delete") -or $employeeID -ne $null -or ($mailNickname -ne $null -or $CSEntryChange.AttributeChanges["mailNickname"].ModificationType -eq "Delete") -or ($targetAddress -ne $null -or  $CSEntryChange.AttributeChanges["targetAddress"].ModificationType -eq "Delete") -or ($businessAreaCode -ne $null -or $CSEntryChange.AttributeChanges["businessAreaCode"].ModificationType -eq "Delete") -or $employeeType -ne $null -or ($orgUnitCode -ne $null -or $CSEntryChange.AttributeChanges["org_unit_cd"].ModificationType -eq "Delete") -or ($SAFlag -ne $null -or $CSEntryChange.AttributeChanges["System_Access_Flag"].ModificationType -eq "Delete") -or ($supervisor_flag -ne $null -or $CSEntryChange.AttributeChanges["supervisor_flag"].ModificationType -eq "Delete") -or ($employeeStatus -ne $null -or $CSEntryChange.AttributeChanges["employeeStatus"].ModificationType -eq "Delete") -or ($personnel_area_cd -ne $null -or $CSEntryChange.AttributeChanges["personnel_area_cd"].ModificationType -eq "Delete") -or ($cost_center_cd -ne $null -or $CSEntryChange.AttributeChanges["cost_center_cd"].ModificationType -eq "Delete") -or ($company_code -ne $null -or $CSEntryChange.AttributeChanges["company_code"].ModificationType -eq "Delete") -or $userType -ne $null)
    if((isModified "exchangeTransitionCode") -or (isModified "deprovisionedDate")-or (isModified "anglczd_first_nm")-or (isModified "anglczd_last_nm") -or (isModified "employeeID") -or (isModified "mailNickname") -or (isModified "targetAddress") -or (isModified "businessAreaCode") -or (isModified "employeeType") -or (isModified "org_unit_cd") -or (isModified "System_Access_Flag") -or (isModified "supervisor_flag") -or (isModified "employeeStatus") -or (isModified "personnel_area_cd") -or (isModified "cost_center_cd") -or (isModified "company_code") -or (isModified "getUserType"))
    {

            $userObject = Get-ADUser $UserObjectID -Properties * -Server $domainController -Credential $Credential
            $mailBoxUserObject = Get-ADUser $UserObjectID -Properties * -Server $domainController -Credential $Credential
            $userHashTable = New-Object HashTable
            $mailBoxPropertiesHashTable = New-Object HashTable
            
			Invoke-UpdateUserCommand $CSEntryChange
            addItemToHashTable $userHashTable "sAMAccountName" $UserObjectID
            addItemToHashTable $mailBoxPropertiesHashTable "sAMAccountName" $UserObjectID

            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
            $null=Set-ADUser -Instance $mailBoxUserObject -Credential $Credential -Server $domainController
            if($exchangeTransitionCode -eq "9")
            {
                if($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_RemoteMailbox)
                {
                    $null=Disable-RemoteMailbox $UserObjectID -DomainController $domainController -confirm:$False
                        
                }
                elseif($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_LocalMailbox)
                {
                        $null=Disable-Mailbox $UserObjectID -DomainController $domainController -confirm:$False
                }
                elseif($userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithSID -or $userObject.msExchRecipientTypeDetails -eq $msExchRecipientTypeDetails_ContactWithoutSID)
                {
				
						 "Export-User-UPDATE: $UserObjectID - Calling diasble mailbox to Clear  mail properties for EA=9, mailcontact to none transition">>$logFileLocation
						 
						 # "Set-ADUser $UserObjectID -clear msExchRecipientTypeDetails -Credential $Credential -Server $domainController">>$logFileLocation
						
						# Contact Release - Code update for Exchange 2016 - Clear msExchRecipientTypeDetails in RF before disabling mailcontact
						#$null=Set-ADUser $UserObjectID -clear msExchRecipientTypeDetails -Credential $Credential -Server $domainController						
						#$null=Set-ADUser -Instance $mailBoxUserObject -Credential $Credential -Server $domainController
						
						
						  "Disable-MailUser $UserObjectID -DomainController $domainController">>$logFileLocation
						
                        $null=Disable-MailUser $UserObjectID -DomainController $domainController -confirm:$False
                }
            }
            
            if($userType -eq "RemoteMailbox" -and $exchangeTransitionCode -eq "1")
            {

                        $RemoteRoutingAddress=$userObject.mailNickname + "@" + $UPNDomain
                        $cmd = "Enable-RemoteMailbox '$UserObjectID' -RemoteRoutingAddress '$RemoteRoutingAddress' -domainController '$domainController'"
                        Invoke-Expression -Command $cmd | Out-Null 
                        
            }
            #elseif($userObject.mailNickname -ne $null -and $userObject.Description -eq $null -and $exchangeFlag -eq "TRUE")
            elseif($userObject.mailNickname -ne $null -and $exchangeFlag -eq "TRUE")
            {
                        $null=Update-Recipient $UserObjectID -DomainController $domainController

            }
            
            set_RecTypeDetails_RecDisplayType -getUserType $getUserType -System_Access_Flag $System_Access_Flag
            $null=Set-ADUser $UserObjectID -Replace $mailBoxPropertiesHashTable -Credential $Credential -Server $domainController
            

            
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


              setADObject_ConnectorAttributes $CSEntryChange
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


Export-CSEntries



Exit-Script -ScriptType "Export" -SuppressErrorCheck -ErrorObject $Error 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 

 
  
 
 
 
 
 
 
 
 