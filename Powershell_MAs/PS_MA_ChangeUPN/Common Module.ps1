  Set-StrictMode -Version 3            
  
  
$Global:ConnectorName = "ExchangeConnector"
$Global:RemoteSessionName = "ExchangeConnector"
$Global:Error.Clear()

          
function New-FIMSchema {            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.Schema])]            
    param()            
            
    [Microsoft.MetadirectoryServices.Schema]::Create()            
}            
            
function New-FIMSchemaType            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.SchemaType])]            
    param            
    (            
        [ValidateNotNullOrEmpty()]            
        [string] $Name,            
        [switch] $LockAnchorAttributeDefinition            
    )            
            
    [Microsoft.MetadirectoryServices.SchemaType]::Create($Name, $LockAnchorAttributeDefinition.ToBool())            
}            
            
function Add-FIMSchemaAttribute            
{            
    [CmdletBinding(DefaultParameterSetName = 'SingleValued')]            
    [OutputType([Microsoft.MetadirectoryServices.SchemaAttribute])]            
    param            
    (            
        [Parameter(Mandatory, ValueFromPipeline)]            
        [ValidateNotNull()]            
        [Microsoft.MetadirectoryServices.SchemaType] $InputObject,            
            
        [Parameter(Mandatory, ParameterSetName='Anchor')]            
        [Parameter(Mandatory, ParameterSetName = 'MultiValued')]            
        [Parameter(Mandatory, ParameterSetName = 'SingleValued')]            
        [ValidateNotNullOrEmpty()]            
        [string] $Name,            
            
        [Parameter(ParameterSetName='Anchor')]            
        [switch] $Anchor,            
            
        [Parameter(ParameterSetName = 'MultiValued')]            
        [switch] $Multivalued,            
            
        [Parameter(Mandatory, ParameterSetName='Anchor')]            
        [Parameter(Mandatory, ParameterSetName = 'MultiValued')]            
        [Parameter(Mandatory, ParameterSetName = 'SingleValued')]            
        [ValidateSet('Binary', 'Boolean', 'Integer', 'Reference', 'String')]            
        [string] $DataType,            
            
        [Parameter(Mandatory, ParameterSetName='Anchor')]            
        [Parameter(Mandatory, ParameterSetName = 'MultiValued')]            
        [Parameter(Mandatory, ParameterSetName = 'SingleValued')]            
        [ValidateSet('ImportOnly', 'ExportOnly', 'ImportExport')]            
        [string] $SupportedOperation            
    )            
                
    switch ($PSCmdlet.ParameterSetName) {            
        'SingleValued' {            
            $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))            
        }            
            
        'MultiValued' {            
            if ($Multivalued) {            
                $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($Name, $DataType, $SupportedOperation))            
            } else {            
                $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))            
            }            
        }            
            
        'Anchor' {            
            if ($Anchor) {            
                $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateAnchorAttribute($Name, $DataType, $SupportedOperation))            
            } else {            
                $InputObject.Attributes.Add([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($Name, $DataType, $SupportedOperation))            
            }            
        }            
    }            
}            
            
function New-FIMCSEntryChange            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.CSEntryChange])]            
    param            
    (            
        [Parameter(Mandatory)]            
        [ValidateNotNullOrEmpty()]            
        [string] $ObjectType,            
            
        [Parameter(Mandatory)]            
        [ValidateSet('Add', 'Delete', 'Update', 'Replace', 'None')]            
        [string] $ModificationType,            
            
        [ValidateNotNullOrEmpty()]            
        [Alias('DistinguishedName')]            
        [string] $DN            
    )            
            
    $CSEntry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()            
    $CSEntry.ObjectModificationType = $ModificationType            
    $CSEntry.ObjectType = $ObjectType            
            
    if ($DN) {            
        $CSEntry.DN = $DN            
    }            
            
    $CSEntry            
}            
            
function Add-FIMCSAttributeChange            
{            
    [CmdletBinding(DefaultParameterSetName = 'Replace')]            
    param            
    (            
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Add')]            
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Update')]            
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Delete')]            
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Replace')]            
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Rename')]            
        [ValidateNotNull()]            
        [Microsoft.MetadirectoryServices.CSEntryChange] $InputObject,            
            
        [Parameter(Mandatory, ParameterSetName='Add')]            
        [switch] $Add,            
            
        [Parameter(Mandatory, ParameterSetName='Update')]            
        [switch] $Update,            
            
        [Parameter(Mandatory, ParameterSetName='Delete')]            
        [switch] $Delete,            
            
        [Parameter(Mandatory, ParameterSetName='Replace')]            
        [switch] $Replace,            
            
        [Parameter(Mandatory, ParameterSetName='Rename')]            
        [switch] $Rename,            
                    
        [Parameter(Mandatory, ParameterSetName='Add')]            
        [Parameter(Mandatory, ParameterSetName='Update')]            
        [Parameter(Mandatory, ParameterSetName='Delete')]            
        [Parameter(Mandatory, ParameterSetName='Replace')]            
        [ValidateNotNullOrEmpty()]            
        [string] $Name,            
            
        [Parameter(Mandatory, ParameterSetName='Add')]            
        [Parameter(Mandatory, ParameterSetName='Update')]            
        [Parameter(Mandatory, ParameterSetName='Replace')]            
        [Parameter(Mandatory, ParameterSetName='Rename')]            
        $Value,            
            
        [switch] $PassThru            
    )            
            
    process {            
        switch ($PSCmdlet.ParameterSetName) {            
            'Add' {            
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($Name, $Value))            
            }            
            
            'Update' {            
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeUpdate($Name, $Value))            
            }            
            
            'Delete' {            
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeDelete($Name))            
            }            
            
            'Replace' {            
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeReplace($Name, $Value))            
            }            
            
            'Rename' {            
                $InputObject.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateNewDN($Value))            
            }            
        }            
            
        if ($PassThru) {            
            $InputObject            
        }            
    }            
}            
            
function New-FIMPutExportEntriesResults            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.PutExportEntriesResults])]            
    param            
    (            
        [ValidateNotNullOrEmpty()]            
        [Microsoft.MetadirectoryServices.CSEntryChangeResult[]] $Results            
    )            
            
    if ($Results) {            
        New-Object Microsoft.MetadirectoryServices.PutExportEntriesResults (New-Object System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChangeResult] (,$Results))            
    } else {            
        New-Object Microsoft.MetadirectoryServices.PutExportEntriesResults            
    }            
}            
            
function New-FIMCloseImportConnectionResults            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.CloseImportConnectionResults])]            
    param            
    (            
        [ValidateNotNullOrEmpty()]            
        [string] $CustomData            
    )            
            
    if ($CustomData) {            
        New-Object Microsoft.MetadirectoryServices.CloseImportConnectionResults $CustomData            
    } else {            
        New-Object Microsoft.MetadirectoryServices.CloseImportConnectionResults            
    }            
}            
            
function New-FIMOpenImportConnectionResults            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.OpenImportConnectionResults])]            
    param            
    (            
        [ValidateNotNullOrEmpty()]            
        [string] $CustomData            
    )            
            
    if ($CustomData) {            
        New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults $CustomData            
    } else {            
        New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults            
    }            
}            
            
function New-FIMCSEntryChanges            
{            
    [CmdletBinding()]            
    [OutputType([System.Collections.Generic.List[Microsoft.MetaDirectoryServices.CSEntryChange]])]            
    param()            
            
    New-Object System.Collections.Generic.List[Microsoft.MetaDirectoryServices.CSEntryChange]            
}            
            
function New-FIMGetImportEntriesResults            
{            
    [CmdletBinding()]            
    [OutputType([Microsoft.MetadirectoryServices.GetImportEntriesResults])]            
    param            
    (            
        [string] $CustomData,            
            
        [switch] $MoreToImport,            
            
        [System.Collections.Generic.List[Microsoft.MetaDirectoryServices.CSEntryChange]] $CSEntries            
    )            
            
    if ($CustomData -or $CSEntries -or $MoreToImport) {            
        New-Object Microsoft.MetadirectoryServices.GetImportEntriesResults $CustomData,$MoreToImport.ToBool(),($CSEntries)            
    } else {            
        New-Object Microsoft.MetadirectoryServices.GetImportEntriesResults            
    }            
}            

function Get-ConfigParameter
{

	[CmdletBinding()]
	[OutputType([string])]
	param(
		[parameter(Mandatory = $true)]
		[System.Collections.ObjectModel.KeyedCollection[string,Microsoft.MetadirectoryServices.ConfigParameter]] 
        $ConfigParameters,
		[parameter(Mandatory = $true)]
		[string]
		$ParameterName,
		[parameter(Mandatory = $false)]
		[ValidateSet("RunStep", "Partition", "Global", "Connectivity", "")]
		[string]
		$Scope,
		[parameter(Mandatory = $false)]
		[switch]
		$Encrypted
	)

	process
	{
		$configParameterValue = $null

		if ([string]::IsNullOrEmpty($Scope))
		{
			$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "RunStep" -Encrypted:$Encrypted

			if ([string]::IsNullOrEmpty($configParameterValue))
			{
				$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Partition" -Encrypted:$Encrypted

				if ([string]::IsNullOrEmpty($configParameterValue))
				{
					$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Global" -Encrypted:$Encrypted

					if ([string]::IsNullOrEmpty($configParameterValue))
					{
						$configParameterValue = Get-ConfigParameter -ConfigParameters $ConfigParameters -ParameterName $ParameterName -Scope "Connectivity" -Encrypted:$Encrypted
					}
				}
			}
		}
		elseif ($Scope -eq "RunStep" -or $Scope -eq "Partition" -or $Scope -eq "Global" -or "Connectivity")
		{
			if ($Scope -eq "RunStep" -or $Scope -eq "Partition" -or $Scope -eq "Global")
			{
				$configParameterName = "{0}_{1}" -f $ParameterName, $Scope
			}
			else
			{
				$configParameterName = $ParameterName
			}

			if ($ConfigParameters.Contains($configParameterName))
			{
				if ($Encrypted -ne $true)
				{
					$configParameterValue =  $ConfigParameters[$configParameterName].Value

					if (![string]::IsNullOrEmpty($configParameterValue))
					{
					   $configParameterValue = $configParameterValue.Trim() 
					}

					Write-Verbose ("ConfigParameter: Scope={0}, Name={1}, Value={2}" -f $Scope, $ParameterName, $configParameterValue)
				}
				else
				{
					$configParameterValue =  $ConfigParameters[$configParameterName].SecureValue

					Write-Verbose ("ConfigParameter: Scope={0}, Name={1}, Value={2}" -f $Scope, $ParameterName, "***Encrypted***")
				}
			}
		}
		else
		{
			throw "Invalid ConfigurationParameter scope: $Scope"
		}

		return $configParameterValue
	}
}

function New-CSEntryChangeExportError
{
	<#
	.Synopsis
		Creates a new CSEntryChangeResult object for the specified CSEntryChange and specified error.
	.Description
		Creates a new CSEntryChangeResult object for the specified CSEntryChange and specified error.
	#>
	
	[CmdletBinding()]
	[OutputType([Microsoft.MetadirectoryServices.CSEntryChangeResult])]
	param(
		[parameter(Mandatory = $true)]
		[Guid]
		$CSEntryChangeIdentifier,
		[parameter(Mandatory = $true)]
		[System.Collections.ArrayList]
		$ErrorObject
	)

	$csentryChangeResult = $null
	# Take the first one otherwise you get "An error occurred while enumerating through a collection: Collection was modified; enumeration operation may not execute.."
	# Seems like a bug in Remote PSH 
	$errorDetail = $ErrorObject[0] # | Out-String -ErrorAction SilentlyContinue
	Write-Warning ("CSEntry Identifier: {0}. ErrorCount: {1}. ErrorDetail: {2}" -f $CSEntryChangeIdentifier, $ErrorObject.Count, $errorDetail)
	$csentryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntryChangeIdentifier, $null, "ExportErrorCustomContinueRun", "RUNTIME_EXCEPTION", $errorDetail)
	Write-Warning ("CSEntryChangeResult Identifier: {0}. ErrorCode: {1}. ErrorName: {2}. ErrorDetail: {3}" -f $csentryChangeResult.Identifier, $csentryChangeResult.ErrorCode, $csentryChangeResult.ErrorName, $csentryChangeResult.ErrorDetail)

	$ErrorObject.Clear()

	return $csentryChangeResult
}

function Enter-Script
{
	<#
	.Synopsis
		Writes the Versbose message saying specified script execution started.
	.Description
		Writes the Versbose message saying specified script execution started.
		Also clear the $Error variable.
	#>
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$ScriptType,
		[parameter(Mandatory = $false)]
		[ValidateNotNull()]
		[System.Collections.ArrayList]
		$ErrorObject
	)

	process
	{
		Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Started..."
		if ($ErrorObject)
		{
			$ErrorObject.Clear()
		}
	}
}

function Exit-Script
{
	<#
	.Synopsis
		Checks $Error variable for any Errors. Writes the Versbose message saying specified script execution sucessfully completed.
	.Description
		Checks $Error variable for any Errors. Writes the Versbose message saying specified script execution sucessfully completed.
		Throws an exception if $Error is present
	#>
	[CmdletBinding()]
	param(
		[parameter(Mandatory = $true)]
		[string]
		$ScriptType,
		[parameter(Mandatory = $false)]
		[ValidateNotNull()]
		[System.Collections.ArrayList]
		$ErrorObject,
		[parameter(Mandatory = $false)]
		[switch]
		$SuppressErrorCheck,
		[parameter(Mandatory = $false)]
		[Type]
		$ExceptionRaisedOnErrorCheck
	)

	process
	{
		if (!$SuppressErrorCheck -and $ErrorObject -and $ErrorObject.Count -ne 0)
		{
			# Take the first one otherwise you get "An error occurred while enumerating through a collection: Collection was modified; enumeration operation may not execute.."
			# Seems like a bug in Remote PSH 
			$errorMessage = $ErrorObject[0] # | Out-String -ErrorAction SilentlyContinue

			if ($ExceptionRaisedOnErrorCheck -eq $null)
			{
				$ExceptionRaisedOnErrorCheck = [Microsoft.MetadirectoryServices.ExtensibleExtensionException]
			}

			$ErrorObject.Clear()

			throw  $errorMessage -as $ExceptionRaisedOnErrorCheck
		}

		Write-Verbose "$Global:ConnectorName - $ScriptType Script: Execution Completed."
	}
}

            
Export-ModuleMember -Function * -Verbose:$false -Debug:$False
 



 
