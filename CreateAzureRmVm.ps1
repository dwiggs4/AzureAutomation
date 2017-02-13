# Include Azure RM Function Library
."$PSScriptRoot\AzureRmFunctionalLibrary.ps1"

# Remove all session variables
Remove-Variable * -ErrorAction Ignore

# Test if user is logged into Azure
try
{
    Get-AzureRmContext 
}
catch [InvalidOperationException]
{
    Write-Host 'Please login to Azure.'
    Login-AzureRmAccount | Out-Null
}

# Select appropriate Azure subscription
Write-Host ''
Write-Host 'Please select a subscription for the new VM.'
$objSubscription = SelectAzureRmSubscription

# Collect required input
Write-Host ''
Write-Host 'Please select a region for the new VM.'
$arrRegions = 'East US','West US'
$strLocation = StringPicker($arrRegions)

# Get/create resource group
Write-Host ''
Write-Host 'Please select a resource group for the new VM.'
Write-Host 'The following resource groups exist in this subscription:'
$objResourceGroup = ObjectPicker(Get-AzureRmResourceGroup)
if($objResourceGroup -eq $null)
{
    CreateResourceGroup
}

# Collect required input
Write-Host ''
Write-Host 'Please select the VM type.'
$arrVmTypes = 'Linux','Windows'
$strVmType = StringPicker($arrVmTypes)

function Get-SitePrefix
{
    Write-Host ''
    Write-Host 'The naming standard for ' -NoNewline; Write-Host 'Azure objects' -ForegroundColor Magenta -NoNewline; Write-Host ' is as follows:'
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

[string]$strSitePrefix = Get-SitePrefix

# Get virtual machine hostname
Write-Host ''
Write-Host 'The following virtual machines exist in this subscription:'
ObjectLister(Get-AzureRmVm)
Write-Host ''
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

Do
{
    $boolValidAzureVMName = $true
    $strAzureVMName = Read-Host -Prompt 'Please provide the name of the Azure virtual machine object'
    $strVMHostname = $strAzureVMName.Substring(6)
    Get-AzureRmVM | ForEach-Object {if ($_.Name -eq $strAzureVMName)
        {
            $boolValidAzureVMName = $false
            Write-Host 'Azure virtual machine' $strAzureVMName ' already exists in this subscription. Select an alternate hostname.'
        }
    }
}
Until ($boolValidAzureVMName -eq $true)

# Instantiate NIC array
# NIC Number (1 indexed), NIC Type (new vs. existing), NIC name, NIC resource group name, NIC location, NIC network name, NIC subnet ID, NIC subnet name (for display purposes), Private IP address
$arrNICs = @()

# Get or create primary NIC
Do
{
    Write-Host ''
    $strUseExistingPrimaryNIC = Read-Host -Prompt 'Use an existing NIC? (Y/N)'
}
Until ($strUseExistingPrimaryNIC.ToLower() -eq 'y' -or $strUseExistingPrimaryNIC.ToLower() -eq 'n')

If ($strUseExistingPrimaryNIC.ToLower() -eq 'y')
{
    # Get existing NIC
    Write-Host ''
    Write-Host 'Please select an existing NIC for the new VM.'
    Write-Host 'The following NICs exist in this subscription and resource group which are not assigned to an existing virtual machine:'
    $objPrimaryNIC = ObjectPicker(Get-AzureRmNetworkInterface -ResourceGroupName $objResourceGroup.ResourceGroupName | Where-Object {$_.VirtualMachine -eq $null})
    $arrNICs += ,@(1,"Existing",$objPrimaryNIC.Name,$objPrimaryNIC.ResourceGroupName,$strLocation,(($objPrimaryNIC.IpConfigurations.Subnet.Id).Split("/"))[-3],$objPrimaryNIC.IPConfigurations.Subnet.Id,(($objPrimaryNIC.IpConfigurations.Subnet.Id).Split("/"))[-1],$objPrimaryNIC.IpConfigurations.PrivateIPAddress)
}
Else
{
    # Get virtual network
    Write-Host ''
    Write-Host 'Please select a virtual network for the new VM.'
    Write-Host 'The following virtual networks exist in this subscription:'
    $objPrimaryVirtualNetwork = ObjectPicker(Get-AzureRmVirtualNetwork)
    
    # Get virtual subnet
    Write-Host ''
    Write-Host 'Please select a virtual subnet for the new VM.'
    Write-Host 'The following virtual subnets exist in this virtual network:'
    $intPrimarySubnetIndex = ObjectPicker($objPrimaryVirtualNetwork.Subnets)
    
    # Set NIC Name
    $strPrimaryNICName = $strSitePrefix + 'NIC' + $strVMHostname + '-1'
    
    # Get private IP and DNS domain
    $strPrimarySubnetAddressPrefix = $objPrimaryVirtualNetwork.Subnets[$intPrimarySubnetIndex].AddressPrefix
    Write-Host ''
    Write-Host "The address range for the selected subnet is $strPrimarySubnetAddressPrefix"
    $strPrimaryPrivateIPAddress = Read-Host -Prompt 'Provide the Virtual Machine private IP Address (xxx.xxx.xxx.xxx format)'
    
    do
    {
        $addPublicIp = Read-Host -Prompt 'Would you like to add a public IP address to the Virtual Machine? (Y/N)'
    }
    Until ($addPublicIp.ToLower() -eq 'y' -or $addPublicIp.ToLower() -eq 'n')

    If ($addPublicIp -eq 'y')
    {
        Write-Host 'Creating public IP address...'
        $strPublicIpName = "$strSitePrefix" + 'PIP' + "$strPrimaryNICName"
        $objPublicIp = New-AzureRmPublicIpAddress -Name $strPublicIpName `
                                                  -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                  -AllocationMethod Static `
                                                  -Location $strLocation `
                                                  -Force `
                                                  -WarningAction Ignore
    }

    

    # Store NIC values in array
    $arrNICs += ,@(1, "New", $strPrimaryNICName, $objResourceGroup.ResourceGroupName, $strLocation,$objPrimaryVirtualNetwork.Name, $objPrimaryVirtualNetwork.Subnets[$intPrimarySubnetIndex].Id, $objPrimaryVirtualNetwork.Subnets[$intPrimarySubnetIndex].Name, $strPrimaryPrivateIPAddress, $objPublicIp.Id)
}

$intNICIndex = 2
# Get additional NIC details
Do
{
    Write-Host ''
    Write-Host 'Larger size VMs can support more than a single NIC.'
    Write-Host '2 NICs: A3, A6, A8, A10, D2, D11, DS2, DS11, G2'
    Write-Host '4 NICs: A4, A7, A9, A11, D3, D12, DS3, DS12, G3'
    Write-Host '8 NICs: D4, D13, D14, DS4, DS13, DS14, G4, G5'
    Write-Host ''
    $strAddAnotherNIC = Read-Host -Prompt 'Add an additional NIC to the new VM? (Y/N)'
    $strAddAnotherNIC = $strAddAnotherNIC.ToLower()
    
    If ($strAddAnotherNIC -eq 'y')
    {
        Do
        {
            Write-Host ''
            $strUseExistingNIC = Read-Host -Prompt 'Use an existing NIC? (Y/N)'
        }
        Until ($strUseExistingNIC.ToLower() -eq 'y' -or $strUseExistingNIC.ToLower() -eq 'n')

        If ($strUseExistingNIC.ToLower() -eq 'y')
        {
            # Get existing NIC
            Write-Host ''
            Write-Host 'Please select an existing NIC for the new VM.'
            Write-Host 'The following NICs exist in this subscription and resource group which are not assigned to an existing virtual machine:'
            $objNIC = ObjectPicker(Get-AzureRmNetworkInterface -ResourceGroupName $objResourceGroup.ResourceGroupName | Where-Object {$_.VirtualMachine -eq $null})
            $arrNICs += ,@($intNICIndex,"Existing",$objNIC.Name,$objNIC.ResourceGroupName,$strLocation,(($objNIC.IpConfigurations.Subnet.Id).Split("/"))[-3],$objNIC.IPConfigurations.Subnet.Id,(($objNIC.IpConfigurations.Subnet.Id).Split("/"))[-1],$objNIC.IpConfigurations.PrivateIPAddress)
            $intNICIndex++
        }
        Else
        {
            # Get virtual network
            Write-Host ''
            Write-Host 'Please select a virtual network for the new VM.'
            Write-Host 'The following virtual networks exist in this subscription:'
            $objVirtualNetwork = ObjectPicker(Get-AzureRmVirtualNetwork)
    
            # Get virtual subnet
            Write-Host ''
            Write-Host 'Please select a virtual subnet for the new VM.'
            Write-Host 'The following virtual subnets exist in this virtual network:'
    
            # Write object list to screen. Cannot use StringPicker function because subnet index is zero indexed.         
            $intSubnetIndex = ObjectPicker($objVirtualNetwork.Subnets)           
            
            <#
            $intNumber = 0
            Foreach ($objVirtualSubnet in $objVirtualNetwork.Subnets)
            {
                Write-Host $intNumber":" $objVirtualSubnet.Name
                $intNumber++
            }
            Write-Host
            $intSubnetIndex = Read-Host -Prompt 'Enter the number of the desired object'
            #>

           # Set NIC Name
           $strPrimaryNICName = $strSitePrefix + 'NIC' + $strVMHostname + '-1'
    
           # Get private IP and DNS domain
           $strPrimarySubnetAddressPrefix = $objPrimaryVirtualNetwork.Subnets[$intPrimarySubnetIndex].AddressPrefix
           Write-Host "The address range for the selected subnet is $strPrimarySubnetAddressPrefix"
           $strPrimaryPrivateIPAddress = Read-Host -Prompt 'Provide the NIC private IP Address (xxx.xxx.xxx.xxx format)'
           $addPublicIp = Read-Host -Prompt 'Would you like to add a public IP address to the NIC? (Y/N)'

           If ($addPublicIp -eq 'y')
           {
               $strPublicIpName = "$strSitePrefix" + 'PIP' + "$strPrimaryNICName"
               $objPublicIp = New-AzureRmPublicIpAddress -Name $strPublicIpName `
                                                         -ResourceGroupName $objResourceGroup `
                                                         -AllocationMethod Static `
                                                         -Location $strLocation `
                                                         -Force `
                                                         -WarningAction Ignore
            }

            # Store NIC values in array
            $arrNICs += ,@($intNICIndex,"New",$strNICName,$objResourceGroup.ResourceGroupName,$strLocation,$objVirtualNetwork.Name,$objVirtualNetwork.Subnets[$intSubnetIndex].Id,$objVirtualNetwork.Subnets[$intPrimarySubnetIndex].Name,$strPrivateIPAddress, $objPublicIp.Id)
            $intNICIndex++
        }
    }
}
Until ($strAddAnotherNIC -ne 'y')

# Set Availability Set
Write-Host ''
$strAddToAvailabilitySet = Read-Host -Prompt 'Add this VM to an availability set (Y/N)'
$strAddToAvailabilitySet = $strAddToAvailabilitySet.ToLower()
If ($strAddToAvailabilitySet -eq 'y')
{
    Write-Host ''
    Write-Host 'Please select an availability set for the new VM.'
    Write-Host 'The following availability sets are available in the same resource group as the new VM:'
    $objAvailabilitySet = ObjectPicker(Get-AzureRmAvailabilitySet -ResourceGroupName $objResourceGroup.ResourceGroupName)
}

# Specify size
Write-Host ''
Write-Host 'Please select a size for the new VM.'
Write-Host 'The following sizes are available:'
$arrVMsizes = Get-AzureRmVMSize -Location $strLocation | Select-Object -ExpandProperty Name
$strVMSize = StringPicker($arrVMSizes)

# Get OS Disk details
Do
{
    Write-Host ''
    $strUseExistingOSDisk = Read-Host -Prompt 'Use an existing OS disk? (Y/N)'
}
Until ($strUseExistingOSDisk.ToLower() -eq 'y' -or $strUseExistingOSDisk.ToLower() -eq 'n')

If ($strUseExistingOSDisk.ToLower() -eq 'y')
{
    # Get existing OSDisk
    $strOSDiskName = 'OSDisk'
    $strOSDiskURI = Read-Host -Prompt 'Provide the URI of the existing OS disk to attach (copy/paste from the Azure web portal)'
}
Else
{
    # Get/create storage account
    Write-Host ''
    Write-Host 'Please select a storage account for the OS disk of new VM.'
    Write-Host 'The following storage accounts exist in this subscription:'
    $objStorageAccount = ObjectPicker(Get-AzureRmStorageAccount)
    
    # Specify publisher, offer, and SKU of image to use
    Write-Host ''
    Write-Host 'Please select a Publisher for the new VM image.'
    Write-Host 'The following publishers are available:'
    
    if($strVmType -like "Windows")
    {
    $arrPublisherNames = 'MicrosoftWindowsServer','MicrosoftSQLServer'
    $strPublisherName = StringPicker($arrPublisherNames)
    }

    elseif($strVmType -like "Linux")
    {
        $arrPublisherNames = 'Canonical','CoreOS', 'credativ', 'OpenLogic', 'RedHat', 'SUSE'
        $strPublisherName = StringPicker($arrPublisherNames)
    }

    Write-Host ''
    Write-Host 'Please select an offer for the new VM image.'
    Write-Host 'For the selected publisher, the following offers are available:'
    $objOffer = ObjectPicker(Get-AzureRmVMImageOffer -Location $strLocation -Publisher $strPublisherName)
    $strOfferName = $objOffer.Offer
    
    Write-Host ''
    Write-Host 'Please select a SKU for the new VM image.'
    Write-Host 'For the selected publisher and offer, the following SKUs are available:'
    $objSKU = ObjectPicker(Get-AzureRmVMImageSKU -Location $strLocation -Publisher $strPublisherName -Offer $strOfferName)
    $strSKUName = $objSKU.Skus
    
    $objLocalAdminCredential = Get-Credential –Message "Please enter the desired local administrator password." `
                                              -UserName ("adm."+($strVMHostname.ToUpper()))
    
    $arrTimeZone = 'Pacific Standard Time', 'Mountaint Standard Time', 'Central Standard Time', 'Eastern Standard Time'
    $strTimeZone = StringPicker($arrTimeZone)

    # Specify OS disk details
    $strOSDiskName = 'OSDisk'
    $strOSDiskURI = $objStorageAccount.PrimaryEndpoints.Blob.ToString() + "vhds/" + $strAzureVMName + "-" + $strOSDiskName + ".vhd"
}

# Instantiate disk array
# LUN number (0 indexed), Disk type (new or existing), Disk size (in GB, $null for no change), Disk name, Storage account name (if known), Storage account object (if possible), Disk URI, CreateOption (empty, attach, fromimage)

$arrAdditionalDataDisks = @()
$intLUNNumber = 0

Do
{
    Write-Host ''
    $strAddAnotherDisk = Read-Host -Prompt 'Add an additional data disk to the new VM? (Y/N)'
    $strAddAnotherDisk = $strAddAnotherDisk.ToLower()
    If ($strAddAnotherDisk -eq 'y')
    {
        Write-Host ''
        $strAddExistingDisk = Read-Host -Prompt 'Is this an existing disk? (Y/N)'
        $strAddExistingDisk = $strAddExistingDisk.ToLower()
        If ($strAddExistingDisk -eq 'y')
        {
            Write-Host ''
            $strDiskName = Read-Host -Prompt 'Provide a name for the existing data disk to attach (for example, "ServerName-Data1.vhd" has a name of "Data1")'
            $strDiskURI = Read-Host -Prompt 'Provide the URI of the existing OS disk to attach (copy/paste from the Azure web portal)'
            $strStorageAccountName = ($strDiskURI.TrimStart("https://").Split("."))[0]
            
            $arrAdditionalDataDisks += ,@($intLUNNumber,'Existing',$null,$strDiskName,$strStorageAccountName,$null,$strDiskURI,'attach')
            $intLUNNumber++
        }
        Else
        {
            Do
            {
                Write-Host ''
                $intDiskSize = Read-Host -Prompt "Please provide disk size (in GB, between 1 and 1023, inclusive) of data disk $intDiskNumber"
                [int]$intDiskSize = [convert]::ToInt32($intDiskSize,10)
    
                If (($intDiskSize -lt 1) -or ($intDiskSize -gt 1023)) 
                {
                    $boolValidDiskSize = $False
                    Write-Host 'Invalid disk size! Disk size must be between 1 and 1023, inclusive.'
                }
                Else
                {
                    $boolValidDiskSize = $True
                }
            }
            Until ($boolValidDiskSize -eq $True)

            Write-Host ''
            $strDiskName = Read-Host -Prompt 'Provide a name for the data disk (e.g. Data1, Data2, Log1, Scratch1, etc)'
            Write-Host ''
            Write-Host 'Please select a storage account for the new data disk.'
            Write-Host 'The following storage accounts exist in this subscription:'
            $objStorageAccount = ObjectPicker(Get-AzureRmStorageAccount)
            $strDiskURI = $objStorageAccount.PrimaryEndpoints.Blob.ToString() + "vhds/" + $strAzureVMName + "-" + $strDiskName + ".vhd"

            $arrAdditionalDataDisks += ,@($intLUNNumber,'New',$intDiskSize,$strDiskName,$objStorageAccount.StorageAccountName,$objStorageAccount,$strDiskURI,'empty')
            $intLUNNumber++
        }
    }
}
Until ($strAddAnotherDisk -ne 'y')


# Get credentials for joining domain
Do
{
    Write-Host ''
    $strJoinDomain = Read-Host -Prompt 'Join this computer to a domain? (Y/N)'
}
Until ($strJoinDomain.ToLower() -eq 'y' -or $strJoinDomain.ToLower() -eq 'n')
    
If ($strJoinDomain.ToLower() -eq 'y')
{
    Write-Host ''
    $strWindowsDomain = Read-Host 'Please provide an Active Directory domain for the new VM'
    $objDomainJoinCredential = Get-Credential -Message ("Please enter the password of an account with permissions to join the "+$strWindowsDomain+" domain") `
                                              -UserName ("svc.domainjoin@"+($strWindowsDomain))
}

Do
{
    Write-Host ''
    $strApplyTag = Read-Host -Prompt 'Tag this VM? (Y/N)'
}
Until ($strApplyTag.ToLower() -eq 'y' -or $strApplyTag.ToLower() -eq 'n')

If ($strApplyTag.ToLower() -eq 'y')
{
    $arrVMTag = SelectTag(ParseTags)
}

# Confirm VM Details
Write-Host ''
Write-Host 'The following VM will be created:'
Write-Host '    Subscription Name: '$objSubscription.SubscriptionName
Write-Host '    Location: '$strLocation
Write-Host '    Site Prefix: '$strSitePrefix
Write-Host '    Resource Group: '$objResourceGroup.ResourceGroupName
Write-Host '    Azure Virtual Machine Name: '$strAzureVMName
Write-Host '    Windows Virtual Machine Hostname: '$strVMHostname

Foreach ($arrNIC in $arrNICs)
{
    Write-Host '    Virtual NIC' $arrNIC[0] 'name: '$arrNIC[2]
    Write-Host '    Virtual NIC' $arrNIC[0] 'type: '$arrNIC[1]
    Write-Host '    Virtual NIC' $arrNIC[0] 'network: '$arrNIC[5]
    Write-Host '    Virtual NIC' $arrNIC[0] 'subnet: '$arrNIC[7]
    Write-Host '    Virtual NIC' $arrNIC[0] 'private IP: '$arrNIC[8]
    
    if($arrNIC[9] -ne $null)
    {
        $pip = Get-AzureRmResource -ResourceId $arrNIC[9]
        Write-Host '    Virtual NIC' $arrNIC[0] 'public IP: '$pip.Properties.ipAddress
    }
}

If ($strAddToAvailabilitySet.ToLower() -eq 'y')
{
    Write-Host '    Availability Set: '$objAvailabilitySet.Name
}
Else
{
    Write-Host '    Availability Set: None'
}

Write-Host '    Virtual Machine Size: '$strVMSize

If ($strUseExistingOSDisk.ToLower() -eq 'y')
{
    Write-Host '    Existing OS Disk URI: '$strOSDiskURI
}
Else
{
    Write-Host '    Virtual Machine Image Publisher: '$strPublisherName
    Write-Host '    Virtual Machine Offer Name: '$strOfferName
    Write-Host '    Virtual Machine SKU: '$strSKUName
    Write-Host '    Virtual Machine Time Zone: '$strTimeZone
    if($strVmType -eq 'Linux')
    {
        $osDiskSize = '30'
    }
    elseif($strVmType -eq 'Windows')
    {
        $osDiskSize - '127'
    }
    Write-Host "    OS Disk Size: $osDiskSize GB (configurable post deployment)"
    Write-Host '    OS Disk Storage Account: '$objStorageAccount.StorageAccountName
    Write-Host '    OS Disk URI: '$strOSDiskURI
}

Foreach ($arrDataDisk in $arrAdditionalDataDisks)
{
    Write-Host "    Data Disk $($arrDataDisk[0]+1) Type: $($arrDataDisk[1])"
    Write-Host "    Data Disk $($arrDataDisk[0]+1) Name: $($arrDataDisk[3])"
    Write-Host "    Data Disk $($arrDataDisk[0]+1) Size: $($arrDataDisk[2])"
    Write-Host "    Data Disk $($arrDataDisk[0]+1) Storage Account: $($arrDataDisk[4])"
    Write-Host "    Data Disk $($arrDataDisk[0]+1) URI: $($arrDataDisk[6])"
}

If ($strJoinDomain.ToLower() -eq 'y')
{
    Write-Host '    Windows Domain Name: '$strWindowsDomain
    Write-Host '    Domain Join Account: '$objDomainJoinCredential.GetNetworkCredential().UserName
}

If ($strApplyTag.ToLower() -eq 'y')
{
    Write-Host "    Tag Name: $($arrVMTag[0])"
    Write-Host "    Tag Value: $($arrVMTag[1])"
}
Write-Host ''

Do
{
    $strCompleteProvisioning = Read-Host -Prompt 'Is everything above correct? (Y/N)'
}
Until ($strCompleteProvisioning.ToLower() -eq 'y' -or $strCompleteProvisioning.ToLower() -eq 'n')

If ($strCompleteProvisioning.ToLower() -eq 'n')
{
    Write-Host 'Exiting script!'
    Exit
}
If ($strCompleteProvisioning.ToLower() -eq 'y')
{
    Write-Host 'Completing VM creation. Please wait...'
    Write-Host ''
    
    # Instantiate VM config
    If ($strAddToAvailabilitySet.ToLower() -eq 'y')
    {
        $objVM = New-AzureRmVMConfig -VMName $strAzureVMName `
                                     -VMSize $strVMSize `
                                     -AvailabilitySetId $objAvailabilitySet.Id
    }
    Else
    {
        $objVM = New-AzureRmVMConfig -VMName $strAzureVMName `
                                     -VMSize $strVMSize
    }
    # Configure options, mount NIC, and mount OS disk

    # Create/Mount NIC(s) as needed
    ForEach ($arrNIC in $arrNICs)
    {
        If ($arrNIC[1] -eq 'Existing')
        {
            If ($arrNIC[0] -eq 1)
            {
                $objNIC = Get-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3]
                $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id -Primary
            }
            Else
            {
                $objNIC = Get-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3]
                $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id
            }
        }
        Elseif ($arrNIC[0] -eq 1 -and $arrNIC[9] -ne $null)
        {
            $objNIC = New-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3] -Location $arrNIC[4] -SubnetId $arrNIC[6] -PrivateIpAddress $arrNIC[8] -PublicIpAddressId $arrNIC[9]
            $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id -Primary
        }
                
        ElseIf ($arrNIC[0] -eq 1 -and $arrNIC[9] -eq $null)
        {
            $objNIC = New-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3] -Location $arrNIC[4] -SubnetId $arrNIC[6] -PrivateIpAddress $arrNIC[8]
            $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id -Primary
        }
            
        ElseIf ($arrNIC[0] -ne 1 -and $arrNIC[9] -ne $null)
        {
            $objNIC = New-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3] -Location $arrNIC[4] -SubnetId $arrNIC[6] -PrivateIpAddress $arrNIC[8] -PublicIpAddressId $arrNIC[9]
            $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id 
        }

        ElseIf ($arrNIC[0] -ne 1 -and $arrNIC[9] -eq $null)
        {
                
            $objNIC = New-AzureRmNetworkInterface -Name $arrNIC[2] -ResourceGroupName $arrNIC[3] -Location $arrNIC[4] -SubnetId $arrNIC[6] -PrivateIpAddress $arrNIC[8]
            $objVM = Add-AzureRmVMNetworkInterface -VM $objVM -Id $objNIC.Id
         }
    }



    # Mount/create OS Disk and other volumes
    If ($strUseExistingOSDisk.ToLower() -eq 'y')
    {
        $objVM = Set-AzureRmVMOSDisk -VM $objVM -Name "OSDisk" -VhdUri $strOSDiskURI -CreateOption attach -Windows
    }
    Else
    {
        if($strVmType -eq 'Windows')
        {
            $objVM = Set-AzureRmVMOperatingSystem -VM $objVM -Windows -ComputerName $strVMHostname -Credential $objLocalAdminCredential -WinRMHttp -ProvisionVMAgent -TimeZone $strTimeZone
        }

        elseif($strVmType -eq 'Linux')
        {
            $objVM = Set-AzureRmVMOperatingSystem -VM $objVM -Linux -ComputerName $strVMHostname -Credential $objLocalAdminCredential
        }
        
        $objVM = Set-AzureRmVMSourceImage -VM $objVM -PublisherName $strPublisherName -Offer $strOfferName -Skus $strSKUName -Version 'latest'
        $objVM = Set-AzureRmVMOSDisk -VM $objVM -Name $strOSDiskName -VhdUri $strOSDiskURI -CreateOption fromImage
    }

    Foreach ($arrDataDisk in $arrAdditionalDataDisks)
    {
        Add-AzureRmVMDataDisk -VM $objVM -Name $arrDataDisk[3] -DiskSizeInGB $arrDataDisk[2] -VhdUri $arrDataDisk[6] -LUN $arrDataDisk[0] -CreateOption $arrDataDisk[7]
    }

    # Create VM
    If ($strApplyTag.ToLower() -eq 'y')
    {
        New-AzureRmVM -ResourceGroupName $objResourceGroup.ResourceGroupName -Location $strLocation -VM $objVM -Tags @(@{Name=$($arrVMTag[0]);Value=$($arrVMTag[1])})
    }
    Else
    {
        New-AzureRmVM -ResourceGroupName $objResourceGroup.ResourceGroupName -Location $strLocation -VM $objVM
    }
        
}
Write-Host 'VM Provisioning complete.'

if (($strUseExistingOSDisk.ToLower() -eq 'n') -and ($strJoinDomain.ToLower() -eq 'y'))
{
    Write-Host ''
    Write-Host 'Joining VM to Domain...'
    Set-AzureRMVMExtension -VMName $strAzureVMName –ResourceGroupName $objResourceGroup.ResourceGroupName -Name "JoinAD" -ExtensionType "JsonADDomainExtension" -Publisher "Microsoft.Compute" -TypeHandlerVersion "1.0" -Location $strLocation -Settings @{ "Name" = $strWindowsDomain; "OUPath" = ""; "User" = "$($objDomainJoinCredential.GetNetworkCredential().UserName)"; "Restart" = "true"; "Options" = 3} -ProtectedSettings @{"Password" = "$($objDomainJoinCredential.GetNetworkCredential().Password)"}
    $objDomainJoinCredential = $null
    Write-Host 'Domain join complete.'
}
Exit
