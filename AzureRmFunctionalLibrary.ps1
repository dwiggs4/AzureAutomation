# This is a 'standard' set of functions to be used and re-used in scripts.
# If you add a function, please document what it does, and maybe where it's useful.

<#
.SYNOPSIS
Displays a list of objects from which the user can select

.DESCRIPTION
Displays a list of objects of a specific type from which the user can select, for use in scripts when 
needing to pick from a list of objects.

.PARAMETER $arrObjects
An array of objects, commonly generated from the appropriate "Get-AzureRM<object type>" commandlet.

.EXAMPLE
ObjectPicker(Get-AzureRMResourceGroup)

Passess the result of "Get-AzureRMResourceGroup" to the function as an array, displays all objects in 
the array, and asks the user to select an object. That object is specified as the return value of the 
function.

ObjectPicker($arrVMs)

Passes "$arrVMs" to the function, displays all objects in that array, and asks the user to select an 
object. That object is specified as the return value of the function.
#>
function ObjectPicker($arrObjects)
{
    # Initialize iterator
    $intNumber = 1

    # Check for empty array
    if ($arrObjects.Count -eq 0)
    {
        Write-Host ''
        Write-Host 'No objects found. One needs to be created.'
        Write-Host ''
        return $null
    }
    else
    {
        # Check for object type
        if ($arrObjects[0].GetType().Name -eq "PSResourceGroup")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.ResourceGroupName)
                $intNumber++
            }
            Write-Host $intNumber":`tCreate and select new Resource Group"
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSAzureSubscription")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.SubscriptionName)
                $intNumber++
            }  
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSStorageAccount")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.StorageAccountName)
                $intNumber++
            }
            Write-Host $intNumber":`tCreate and select new Storage Account"
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSSubnet")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.Name)
                $intNumber++
            }
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSNetworkInterface")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.Name)
                $intNumber++
            }
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSAzureSubscription")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.SubscriptionName)
                $intNumber++
            }
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSVirtualMachineImageOffer")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.Offer)
                $intNumber++
            }
        }
        elseif ($arrObjects[0].GetType().Name -eq "PSVirtualMachineImageSku")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.Skus)
                $intNumber++
            }
        }
        else
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
            {
                Write-Host $intNumber":`t"$($objObject.Name)
                $intNumber++
            }
        }     
    }    
    
    # Prompt for user input    
    if ($arrObjects.Count -gt 0)
    {
        Write-Host ''
        $intSelection = Read-Host -Prompt 'Enter the number of the desired option'
        $intCurrentNumber = $intNumber - 1
        
        # Check to see if new object is required
        if ($intSelection -eq $intNumber -and (($arrObjects[0].GetType().Name -eq 'PSResourceGroup') `
        -or ($arrObjects[0].GetType().Name -eq 'PSStorageAccount') `
        -or ($arrObjects[0].GetType().Name -eq 'PSAvailabilitySet')))
        {
            # A new object is required
            Switch ($arrObjects[0].GetType().Name)
            {
                'PSResourceGroup'
                {
                    return CreateResourceGroup
                }
                'PSStorageAccount'
                {
                    return CreateStorageAccount
                }
                'PSAvailabilitySet'
                {
                    return CreateAvailabilitySet
                }
            }
        }
        # Subnets are zero-indexed so return the selection number minus 1
        elseif ($arrObjects[0].GetType().Name -eq 'PSSubnet')
        {
             return ($intSelection-1)    
        }
        elseif ($intSelection -notin 1..$intCurrentNumber)
        {
            Write-Host ''
            Write-Host 'Error in confirming selection.' -ForegroundColor Red
            Write-Host 'Please try again.' -ForegroundColor Red
            Write-Host ''
            ObjectPicker($arrObjects)
        }        
        else
        {
            return $arrObjects[$intSelection-1]
        }
    }
}

<#
.SYNOPSIS
Displays a list of objects

.DESCRIPTION
Displays a list of objects of a specific type, for use in scripts when needing to display (but not 
pick from) a list of objects.

.PARAMETER $arrObjects
An array of objects, commonly generated from the appropriate "Get-AzureRM<object type>" commandlet.

.EXAMPLE
ObjectLister(Get-AzureRMResourceGroup)

Passes the result of "Get-AzureRMResourceGroup" to the function as an array, and displays all objects 
in the array.

ObjectLister($arrVMs)

Passes "$arrVMs" to the function and displays all objects in that array.
#>
function ObjectLister($arrObjects)
{
    if($arrObjects -ne $null)
    {
        # Initialize iterator
        $intNumber = 1

        # Check for object type
        if ($arrObjects[0].GetType().Name -eq "PSResourceGroup")
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
                {
                    Write-Host $intNumber":`t"$($objObject.ResourceGroupName)
                    $intNumber++
                }
        }
        else
        {
            # Write object list to screen
            foreach ($objObject in $arrObjects)
                {
                    Write-Host $intNumber":`t"$($objObject.Name)
                    $intNumber++
                }
        }     
    }
}

<#
.SYNOPSIS
Displays a list of strings from which the user can select

.DESCRIPTION
Displays a list of strings of from which the user can select, for use in scripts when needing to pick 
from a list of strings.

.PARAMETER $arrObjectsget
An array of strings, commonly user generated.

.EXAMPLE
StringPicker($arrStrings)

Passes "$arrStrings" to the function, displays all strings in that array, and asks the user to select 
a string. That string is specified as the return value of the function.
#>
function StringPicker($arrStrings)
{
    # Initialize iterator
    $intNumber = 1

    # Write string list to screen
    foreach ($objStrings in $arrStrings)
    {
        Write-Host $intNumber":`t"$($objStrings)
        $intNumber++
    }

    # Prompt for user input
    Write-Host
    $intSelection = Read-Host -Prompt 'Enter the number of the desired option'
    
    $intCurrentNumber = $intNumber - 1 
    if ($intSelection -notin 1..$intCurrentNumber)
    {
        Write-Host ''
        Write-Host 'Error in confirming selection.' -ForegroundColor Red
        Write-Host 'Please try again.' -ForegroundColor Red
        Write-Host ''
        StringPicker($arrStrings)
    }
    
    else
    {
        return $arrStrings[$intSelection-1]
    }
}

<#
.SYNOPSIS
Selects an Azure RM subscription

.DESCRIPTION
Displays a list of Azure RM subscriptions to which the currently logged in user has access and prompts 
for selection.

.EXAMPLE
SelectAzureRmSubscription

Displays a list of Azure RM subscriptions to which the currently logged in user has access and prompts 
for selection.
#>
function SelectAzureRmSubscription()
{
    # Pick subscription
    $objSubscription = ObjectPicker(Get-AzureRmSubscription -WarningAction Ignore)

    Select-AzureRmSubscription -SubscriptionId $objSubscription.SubscriptionId
    return $objSubscription
}

<#
.SYNOPSIS
Automates the login process for Azure Resource Manager Mode

.DESCRIPTION
Adds an Azure RM account and selects a subscription from the list of subscriptions for that account

.EXAMPLE
LoginToARM()

Prompts the user to login and then displays a list of subscriptions to which that account has access, 
for selection by the user.
#>
function LoginToARM()
{
    # Initialize iterator
    $intNumber = 1

    # Add Azure Account
    Login-AzureRmAccount

    # Select Azure Subscription
    SelectAzureRmSubscription    
}

<#
.SYNOPSIS
Creates a resource group

.DESCRIPTION
Creates a resource group based on user input

.EXAMPLE
CreateResourceGroup()

Prompts the user for the name and location of a desired resource group to be created.
#>
function CreateResourceGroup()
{
    Write-Host 'The naming standard for resource groups is as follows:'
    Write-Host ''
    Write-Host '        xxx' -ForegroundColor Yellow -NoNewline; Write-Host 'RSG' -ForegroundColor Magenta -NoNewline; Write-Host 'Description' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Where:'
    Write-Host '    xxx' -ForegroundColor Yellow -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription'
    Write-Host '    RSG' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character code for a resource group object'
    Write-Host '    Description' -ForegroundColor Cyan -NoNewline; Write-Host ' = a text description of the objects the resource group will contain'
    Write-Host ''
    Write-Host 'Please provide detail on the resource group to be created:'
    Write-Host ''
    
    # Prompt for input
    $strResourceGroupName = Read-Host -Prompt 'New resource group name'

    Write-Host ''
    Write-Host 'Please select a region for the new resource group.'
    $arrRegions = 'East US','West US'
    $strResourceGroupLocation = StringPicker($arrRegions)
    
    # Create new resource group
    Write-Host 'Creating new resource group...'
    $objNewResourceGroup = New-AzureRmResourceGroup -Name $strResourceGroupName `
                                                    -Location $strResourceGroupLocation
    return $objNewResourceGroup
}

<#
.SYNOPSIS
Creates a storage account

.DESCRIPTION
Creates a storage account based on user input

.EXAMPLE
CreateStorageAccount()

Prompts the user for the name, resource group, storage account type, and location of a desired storage 
account to be created.
#>
function CreateStorageAccount()
{
    Write-Host 'The naming standard for storage accounts is as follows (note that lower case is required):'
    Write-Host ''
    Write-Host '        xxx' -ForegroundColor Yellow -NoNewline; Write-Host 'sto | stp' -ForegroundColor Magenta -NoNewline; Write-Host 'description' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Where:'
    Write-Host '    xxx' -ForegroundColor Yellow -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription'
    Write-Host '    sto | stp' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character code for a standard or premium storage account object, respectively'
    Write-Host '    description' -ForegroundColor Cyan -NoNewline; Write-Host ' = a *LOWERCASE ALPHANUMERIC* text description of the objects the storage account will contain, of no more than 26 characters in length'
    Write-Host ''
    Write-Host 'Please provide detail on the storage account to be created:'
    Write-Host ''

    # Prompt for input
    $strStorageAccountName = Read-Host -Prompt 'New storage account name'
    $strStorageAccountName = $strStorageAccountName.ToLower()
    Write-Host 'Please select a resource group for the new storage account.'
    Write-Host 'The following resource groups exist in this subscription:'
    $objStorageAccountResourceGroup = ObjectPicker(Get-AzureRmResourceGroup)
    $strStorageAccountResourceGroupName = $objStorageAccountResourceGroup.ResourceGroupName
    $arrStorageAccountTypes = 'Standard_LRS','Standard_ZRS','Standard_GRS','Standard_RAGRS','Premium_LRS'
    $strStorageAccountType = StringPicker($arrStorageAccountTypes)

    Write-Host ''
    Write-Host 'Please select a region for the new storage account.'
    $arrRegions = 'East US','West US'
    $strStorageAccountLocation = StringPicker($arrRegions)
    
    # Create new storage account
    Write-Host 'Creating new storage account...'
    $objNewStorageAccount = New-AzureRmStorageAccount -Name $strStorageAccountName `
                                                      -ResourceGroupName $strStorageAccountResourceGroupName `
                                                      -Type $strStorageAccountType `
                                                      -Location $strStorageAccountLocation
    return $objNewStorageAccount
}

<#
.SYNOPSIS
ToDo

.DESCRIPTION
ToDo

.EXAMPLE
ToDo
#>
function GetSitePrefix()
{
    Write-Host ''
    Write-Host 'The naming standard for Azure objects is as follows:'
    Write-Host ''
    Write-Host '        xxx' -ForegroundColor Magenta -NoNewline; Write-Host 'YYYDescription'
    Write-Host ''
    Write-Host 'Where:' 
    Write-Host '    xxx' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription' 
    Write-Host '    YYY = three-character object code identifying the object type'
    Write-Host '    Description = text string describing the object'
    Write-Host ''
    $sitePrefix = Read-Host -Prompt 'Please provide the site code/uniqe identifier for the new VM'
    $sitePrefix = $sitePrefix.ToLower()
    [regex]$regSitePrefix = '\b[a-zA-Z]{3}\b'
    if($sitePrefix -notmatch $regSitePrefix)
    {
        Write-Host ''
        Write-Host 'Error in confirming the site code/unique identifyer.' -ForegroundColor Red
        Write-Host 'Please try again.' -ForegroundColor Red
        Get-SitePrefix
    }
    return $sitePrefix
}

<#
.SYNOPSIS
ToDo

.DESCRIPTION
ToDo

.EXAMPLE
ToDo
#>
function GetVmHostname($strVmType)
{
    if($strVmType -like "Linux")
    {
        Write-Host ''
        Write-Host 'The naming standard for Linux virtual machines is as follows:'
        Write-Host ''
        Write-Host '        xxx' -ForegroundColor Yellow -NoNewline; Write-Host 'LVM' -ForegroundColor Magenta -NoNewline; Write-Host 'Hostname' -ForegroundColor Cyan
        Write-Host ''
        Write-Host 'Where:'
        Write-Host '    xxx' -ForegroundColor Yellow -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription' 
        Write-Host '    LVM' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character code for a Linux virtual machine object' 
        Write-Host '    Hostname' -ForegroundColor Cyan -NoNewline; Write-Host ' = the Linux hostname of the new virtual machine' 
        Write-Host ''
    }

    elseif($strVmType -like "Windows")
    {
        Write-Host 'The naming standard for Windows virtual machines is as follows:'
        Write-Host ''
        Write-Host '        xxx' -ForegroundColor Yellow -NoNewline; Write-Host 'WVM' -ForegroundColor Magenta -NoNewline; Write-Host 'Hostname' -ForegroundColor Cyan
        Write-Host ''
        Write-Host 'Where:'
        Write-Host '    xxx' -ForegroundColor Yellow -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription'
        Write-Host '    LVM' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character code for a Windows virtual machine object'
        Write-Host '    Hostname' -ForegroundColor Cyan -NoNewline; Write-Host ' = the Windows hostname of the new virtual machine'
        Write-Host ''
    }

    do
    {
        $boolValidAzureVMName = $true
        $strAzureVMName = Read-Host -Prompt 'Please provide the name of the Azure virtual machine object'
        Get-AzureRmVM | `
        ForEach-Object `
        {
            if ($_.Name -eq $strAzureVMName)
            {
                $boolValidAzureVMName = $false
                Write-Host 'Azure virtual machine' $strAzureVMName ' already exists in this subscription. Select an alternate hostname.'
            }
        }
    }
    until ($boolValidAzureVMName -eq $true)
    return $strAzureVMName
}



<#
.SYNOPSIS
Creates an availability set

.DESCRIPTION
Creates an availability set based on user input

.EXAMPLE
CreateAvailabilitySet()

Prompts the user for the name, resource group, and location of a desired availablity set to be created.
#>
function CreateAvailabilitySet()
{
    Write-Host 'The naming standard for availability sets is as follows:'
    Write-Host ''
    Write-Host '        xxx' -ForegroundColor Yellow -NoNewline; Write-Host 'ASV' -ForegroundColor Magenta -NoNewline; Write-Host 'Description' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Where:'
    Write-Host '    xxx' -ForegroundColor Yellow -NoNewline; Write-Host ' = the three-character site code/unique identifier for the current Azure subscription'
    Write-Host '    AVS' -ForegroundColor Magenta -NoNewline; Write-Host ' = the three-character code for an availability set object'
    Write-Host '    Description' -ForegroundColor Cyan -NoNewline; Write-Host ' = a text description of the objects the availability set will contain'
    Write-Host ''
    Write-Host 'Please provide detail on the availability set to be created:'
    Write-Host ''

    # Prompt for input
    $strAvailabilitySetName = Read-Host -Prompt 'New availability set name'
    Write-Host ''
    Write-Host 'Please select a resource group for the new availability set.'
    Write-Host 'The following resource groups exist in this subscription:'
    $objAvailabilitySetResourceGroup = ObjectPicker(Get-AzureRmResourceGroup)
    $strAvailabilitySetResourceGroupName = $objAvailabilitySetResourceGroup.ResourceGroupName

    Write-Host ''
    Write-Host 'Please select a region for the new availabilty set.'
    $arrRegions = 'East US','West US'
    $strAvailabilitySetLocation = StringPicker($arrRegions)

    $intPlatformFaultdomainCount = 2
    $intPlatformUpdatedomainCount = 5

    # Create new availability set
    Write-Host 'Creating new availability set...'
    $objNewAvailabilitySet = New-AzureRmAvailabilitySet -Name $strAvailabilitySetName `
                                                        -ResourceGroupName $strAvailabilitySetResourceGroupName `
                                                        -Location $strAvailabilitySetLocation `
                                                        -PlatformUpdatedomainCount $intPlatformUpdatedomainCount `
                                                        -PlatformFaultdomainCount $intPlatformFaultdomainCount | Out-Null
    return $objNewAvailabilitySet
}

<#
.SYNOPSIS
Parses all available Azure tags in to a structured array

.DESCRIPTION
Parses the output of the Get-AzureRmTag -Detailed cmdlet and stores results in a structured 2-column 
array. Column 0 is the Key/Tag Name, Column 1 is the Value/Tag Value.

.EXAMPLE
ParseTags()

Returns an array with structured tag data
#>
function ParseTags()
{
    # Get and convert the output of the Get-AzureRMTag command so it can be parsed
    Get-AzureRmTag -Detailed | Out-File "C:\Tags.txt"
    $arrTagText = Get-Content "C:\Tags.txt"
    
    # Define an empty array to fill with tag data
    $arrTags = @()
    
    foreach ($strTagTextLine in $arrTagText)
    {
        if ($strTagTextLine.ToLower().StartsWith('name'))
        {
            # Check for the "Name" header indicating a group of tags
            $strCurrentTagName = $strTagTextLine.SubString($strTagTextLine.IndexOf(":") + 2)
            Continue
        }
        if ($strTagTextLine.ToLower().StartsWith("         n"))
        {
            # Determine the length of the Value field (confusingly also listed as "Name" in the 
            # Get-AzureRmTag output)
            $intPositionOfCount = $strTagTextLine.IndexOf("Count")
            Continue
        }
        if ($strTagTextLine.ToLower().StartsWith("         ="))
        {
            # Check for the separator line indicating that tag values are about to be listed
            $boolExpectValues = $true
            Continue
        }
        if (($boolExpectValues) -and ($strTagTextLine.Length -gt 9))
        {
            # Parse for tag values and store to an array
    
            # Strip tag counts
            $strTemp = $strTagTextLine.Substring(0,$intPositionOfCount)
            # Strip spaces
            $strTemp = $strTemp.Trim()
            # Write to array
            $arrTags += ,@($strCurrentTagName,$strTemp)
            Continue
        }
        if (($boolExpectValues) -and ($strTagTextLine.Length -le 9))
        {
            # Check to see if the list of values for a given name is done
            $boolExpectValues = $false
            Continue
        }
    }
    $arrTags = $arrTags | Sort-Object
    return $arrTags
}

<#
.SYNOPSIS
Displays a list of tags from which the user can select

.DESCRIPTION
Displays a list of tags from which the user can select, for use in scripts when needing to pick from 
a list of objects.

.PARAMETER $arrTags
An 2 dimensional array of tags, commonly generated from the ParseTags() function. Column 0 must be the 
Key/Tag Name, Column 1 must be the Value/Tag Value.

.EXAMPLE
SelectTag(ParseTags())

Passess the result of "ParseTags()" to the function as an array, displays all tags in the array, and 
asks the user to select a tag. That tag is the return value of the function.
#>
function SelectTag($arrTags)
{
    # Initialize iterator
    $intNumber = 1

    # Check for empty array
    if ($arrTags.Count -eq 0)
    {
        Write-Host 'No objects of that type found.'
        Write-Host $intNumber":`tCreate and select new Tag"
    }
    else
    {
        # Write object list to screen
        foreach ($arrTag in $arrTags)
        {
            Write-Host $intNumber":`tName:" $($arrTag[0])"`tValue:" $($arrTag[1])
            $intNumber++
        }
        Write-Host $intNumber":`tCreate and select new Tag"
    }

    # Prompt for user input
    Write-Host
    $intSelection = Read-Host -Prompt 'Enter the number of the desired option'

    # Check to see if new object is required
    if ($intSelection -eq $intNumber)
    {
        return CreateTag
    }
    else
    {
        return $arrTags[$intSelection-1]
    }
}

<#
.SYNOPSIS
Creates an Azure tag

.DESCRIPTION
Creates an Azure tag based on user input

.EXAMPLE
CreateTag()

Prompts the user for the tag name, and tag value, creates the tag, and returns an array storing the 
name and value.
#>
function CreateTag()
{
    # Prompt for input
    $strTagName = Read-Host -Prompt 'New tag name'
    $strTagValue = Read-Host -Prompt 'New tag value'

    # Create new availability tag
    $objTag = New-AzureRmTag -Name $strTagName -Value $strTagValue
    
    # return array for specific tag created, as the New-AzureRMTag cmdlt returns detail on all Values 
    # associated with a given Name
    return @($strTagName,$strTagValue) 
}
