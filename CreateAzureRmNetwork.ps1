# Remove variables that may be residual in the session
Remove-Variable * -ErrorAction Ignore

# Script requires AzureRmNetwork.xlsx
if(-not (Test-Path .\AzureRmNetwork.xlsx))
{
    (Get-Date -Format MM/dd/yyy_hh:mm:ss) + " ERROR: Cannot find helper file AzureRmNetwork.xlsx. Please ensure the working directory includes the Excel Workbook. Ending script."
    exit
}

# Include Azure Rm Function Library
if(Test-Path .\AzureRmFunctionalLibrary.ps1)
{
    . ".\AzureRmFunctionalLibrary.ps1"
}

else 
{
    (Get-Date -Format MM/dd/yyy_hh:mm:ss) + " ERROR: Cannot find helper file AzureRmFunctionalLibrary.ps1. Please ensure the working directory includes the Excel Workbook. Ending script."
    exit
}

# Include Excel Function Library
if(Test-Path .\ExcelFunctionalLibrary.ps1)
{
    . ".\ExcelFunctionalLibrary.ps1"
}

else 
{
    (Get-Date -Format MM/dd/yyy_hh:mm:ss) + " ERROR: Cannot find helper file ExcelFunctionalLibrary.ps1. Please ensure the working directory includes the Excel Workbook. Ending script."
    exit
}

# Login to Azure with a resource manager account
Login-AzureRmAccount | Out-Null
$account = (Get-AzureRmContext | select Account -ExpandProperty Account)
(Get-Date -Format MM/dd/yyy_hh:mm:ss) + " INFO: Using account, $account."

# Get current subscriptions 
[array]$subscriptions = Get-AzureRmSubscription -WarningAction Ignore

if ($subscriptions.Count -gt 1)
{
    Write-Output ''
    Write-Output 'Please select a subscription for the network update.'

    # Select appropriate Azure subscription
    # Function SelectAzureRmSubscription included in AzureRmFunctionalLibrary
    $subscription = SelectAzureRmSubscription
    $subscriptionName = $subscription.Subscription.SubscriptionName
    (Get-Date -Format MM/dd/yyy_hh:mm:ss) + " INFO: Using subscription, $subscriptionName."
}

else
{
    $subscriptionId = $subscriptions[0].SubscriptionId
    $subscription = Select-AzureRmSubscription -SubscriptionId $subscriptionId
    $subscriptionName = $subscription.Subscription.SubscriptionName
    (Get-Date -Format MM/dd/yyy_hh:mm:ss) + " INFO: Using subscription, $subscriptionName."
}

# Get resource groups to be used
# Import-Excel function included in AzureRmFunctionalLibrary
[array]$resourceGroups = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                      -SheetName "Resource Groups" `
                                      -closeExcel

# Create listed resource groups if the object is not null
# resourceGroups properties pulled from Import-Excel function
if($resourceGroups -ne $null)
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: Will now create the resource group(s)."
    Write-Output $output
        
    # Get current resource groups if any
    $currentResourceGroups = Get-AzureRmResourceGroup
    $currentResourceGroups | % {[array]$currentResourceGroupNames += $_.ResourceGroupName}
    
    foreach ($resourceGroup in $resourceGroups)
    {
        $resourceGroupName = $resourceGroup.resourceGroupName
        [array]$resourceGroupNames += $resourceGroupName
        Write-Progress -Activity "Creating resource groups.." `
                       -Status "Working on $resourceGroupName" `
                       -PercentComplete ((($resourceGroups.IndexOf($resourceGroup)) / $resourceGroups.Count) * 100)
    
        # Check to see if resource group already exists
        if ($resourceGroupName -inotin $currentResourceGroupNames)
        {
            New-AzureRmResourceGroup -Name $resourceGroupName `
                                     -Location $resourceGroup.location `
                                     -Force | Out-Null
        } 

        else
        {
            Write-Output ''
            $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
            $output = $time + " INFO: The resource group, $resourceGroupName, already exists."
            Write-Output $output
        }
}

Write-Progress -Activity "Creating resource groups.." `
               -Status "Done" `
               -PercentComplete 100 `
               -Completed
}

else
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: Did not find any resource groups to create."
    Write-Output $output
}

# Reconcile what resource groups exist that were not included in the helper document AzureRmNetwork.xlsx
$extraResourceGroups = $currentResourceGroupNames | ? {$resourceGroupNames -NotContains $_}

if ($extraResourceGroups -ne $null)
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: The following resource groups were found in the subscription, but were not found in the AzureRmNetwork.xlsx document."
    Write-Output $output
    Write-Output ''
    $output = $extraResourceGroups
    Write-Output $output
}

# Get storage accounts to be used
# Import-Excel function included in AzureRmFunctionalLibrary
[array]$storageAccounts = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                       -SheetName "Storage Accounts" `
                                       -closeExcel

# Create listed storage accounts if the object is not null
# storageAccounts properties are pulled from the Import-Excel function
# Don't forget to enable the 'Secure transfer required' option in the azure portal to enable access
# via HTTPS, strictly
# As of 7/7/2017 this option is not available to set via PowerShell
if($storageAccounts -ne $null)
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: Will now create the storage account(s)."
    Write-Output $output

    # Get current storage accounts if any
    $currentStorageAccounts = Get-AzureRmStorageAccount
    $currentStorageAccounts | % {[array]$currentStorageAccountNames += $_.StorageAccountName}

    foreach ($storageAccount in $storageAccounts)
    {
        $storageAccountName = $storageAccount.storageAccountName
        $storageAccountNames += $storageAccountName
        Write-Progress -Activity "Creating storage accounts.." `
                       -Status "Working on $storageAccountName" `
                       -PercentComplete ((($storageAccounts.IndexOf($storageAccount)) / $storageAccounts.Count) * 100)
    
        # Check to see if storage account already exists
        if ($storageAccountName -inotin $currentStorageAccountNames)
        {
            # If kind is like blob storage more properties are required
            if ($storageAccount.kind -like "BlobStorage")
            {
                New-AzureRmStorageAccount -Name $storageAccount.storageAccountName `
                                          -ResourceGroupName $storageAccount.resourceGroupName `
                                          -SkuName $storageAccount.skuName `
                                          -Location $storageAccount.location `
                                          -Kind $storageAccount.kind `
                                          -AccessTier $storageAccount.accessTier `
                                          -EnableEncryptionService $storageAccount.enableEncryptionService | Out-Null
            }

            else
            {
                New-AzureRmStorageAccount -Name $storageAccount.storageAccountName `
                                          -ResourceGroupName $storageAccount.resourceGroupName `
                                          -SkuName $storageAccount.skuName `
                                          -Location $storageAccount.location `
                                          -Kind $storageAccount.kind | Out-Null
            }                              
        }

        else
        {
            Write-Output ''
            $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
            $output = $time + " INFO: The storage account, $storageAccountName, already exists."
            Write-Output $output
        }        
    }
    Write-Progress -Activity "Creating storage accounts.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any storage accounts to create.'
}

# Reconcile what storage accounts exist that were not included in the helper document AzureRmNetwork.xlsx
$extraStorageAccounts = $stoargeAccountNames | ? {$currentStorageAccountNames -NotContains $_}

if ($extraStorageAccounts -ne $null)
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: The following storage accounts were found in the subscription, but were not found in the AzureRmNetwork.xlsx document."
    Write-Output $output
    Write-Output ''
    $output = $extraStorageAccounts
    Write-Output $output
}

# Get subnets to be deployed
# Import-Excel function included in AzureRmFunctionalLibrary
[array]$importSubnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                               -SheetName "Subnets" `
                               -closeExcel

# Possible script enhancement would be to check the syntax of the CIDR for a given subnet

# Get current subnets if any
Get-AzureRmVirtualNetwork | % {[array]$currentSubnets += $_.Subnets}
$currentSubnets = $currentSubnets | select Name, AddressPrefix

$compareSubnets = Compare-Object $currentSubnets $subnets -Property Name, AddressPrefix -IncludeEqual

foreach ($subnet in $compareSubnets)
{
    $vnetName = $subnet.VNET
    $subnetName = $subnet.Name
    $sideIndicator = $subnet.SideIndicator
    
    if ($sideIndicator -like '==')
    {
        Write-Output ''
        $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
        $output = $time + " INFO: The subnet, $subnetName, in VNET, $vnetName, already exists."
        Write-Output $output
    }

    elseif ($sideIndicator -like '<=')
    {
        Write-Output ''
        $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
        $output = $time + " INFO: The subnet, $subnetName, in VNET, $vnetName, exists but is not in the inventory of subnets in AzureRmNetwork.xlsx."
        Write-Output $output
    }
}

if($subnets -ne $null)
{
    [Reflection.Assembly]::LoadWithPartialname('Microsoft.Azure.Commands.Network.Models.PSSubnet') | Out-Null
        
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: Will now create the subnet(s)."
    Write-Output $output
  
    # Create hashtable that groups virtual networks and associated subnets
    $virtualNetworks = @{}

    foreach ($subnet in $subnets)
    {
        $vnetName = $subnet.VNET
        $subnetName = $subnet.Name
        
        if($virtualNetworks.Keys -notcontains $subnet.VNET)
        {
            $subnetList = New-Object 'System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSubnet]'
            $virtualNetworks.Add("$vnetName", $subnetList)
            Write-Progress -Activity "Creating network subnets.." `
                           -Status "Working on $subnetName" `
                           -PercentComplete ((($subnets.IndexOf($subnet)) / $subnets.Count) * 100)
            $subnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name $subnet.Name `
                                                                  -AddressPrefix $subnet.AddressPrefix
            $expression = '$virtualNetworks.' + "$vnetName" + '.Add($subnetConfig)'
            Invoke-Expression $expression
        }

        else
        {
            Write-Progress -Activity "Creating network subnets.." `
                            -Status "Working on $subnetName" `
                            -PercentComplete ((($subnets.IndexOf($subnet)) / $subnets.Count) * 100)
            $subnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name $subnet.Name `
                                                                  -AddressPrefix $subnet.AddressPrefix
            $expression = '$virtualNetworks.' + "$vnetName" + '.Add($subnetConfig)'
            Invoke-Expression $expression
        }
    }
       
Write-Progress -Activity "Creating network subnets.." `
                -Status "Done" `
                -PercentComplete 100 `
                -Completed
}

else
{
    Write-Output ''
    $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
    $output = $time + " INFO: Did not find any subnets to create."
    Write-Output $output
}

# Get VNETs to be deployed
# Import-Excel function included in AzureRmFunctionalLibrary
[array]$vnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                             -SheetName "VNETs" `
                             -closeExcel

if($vnets -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the VNET(s).'

    # Get current VNETs if any
    $currentVnets = Get-AzureRmVirtualNetwork
    $currentVnets | % {[array]$vnetNames += $_.Name}
    
    # Create virtual networks
    foreach ($vnet in $vnets)
    {      
        $vnetName = $vnet.vnetName    
        Write-Progress -Activity "Creating virtual networks.." `
                       -Status "Working on $vnetName" `
                       -PercentComplete ((($vnets.IndexOf($vnet)) / $vnets.Count) * 100)

        if ($vnetName -inotin $vnetNames)
        {
            $expression = '$virtualNetworks.' + "$vnetName"
            $subnetList = (Invoke-Expression $expression)

            New-AzureRmVirtualNetwork -Name $vnet.vnetName `
                                      -ResourceGroupName $vnet.resourceGroupName `
                                      -Location $vnet.location `
                                      -AddressPrefix $vnet.CIDR `
                                      -Subnet $subnetList `
                                      -WarningAction Ignore | Out-Null
        } 

        else
        {
            Write-Output ''
            Write-Output "The virtual network, $vnetName, already exists!"
        }
}

Write-Progress -Activity "Creating virtual networks.." `
               -Status "Done" `
               -PercentComplete 100 `
               -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any VNETs to create.'
}

[array]$networkSecurityRules = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                            -SheetName "Network Security Rules" `
                                            -closeExcel

if($networkSecurityRules -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the network security rule(s).'

    # Initialize empty hash table
    $networkSecurityGroups = @{}

    # Itterate through network security rules object
    foreach ($rule in $networkSecurityRules)
    {
        $ruleName = $rule.Name
        Write-Progress -Activity "Getting network security rules.." `
                       -Status "Working on $ruleName" `
                       -PercentComplete ((($networkSecurityRules.IndexOf($rule)) / $networkSecurityRules.Count) * 100)

        # Create rule config from row of csv
        $ruleConfig = New-AzureRmNetworkSecurityRuleConfig -Name $rule.Name `
                                                           -Access $rule.Access `
                                                           -Protocol $rule.Protocol `
                                                           -Direction $rule.Direction `
                                                           -Priority $rule.Priority `
                                                           -SourceAddressPrefix $rule.SourceAddressPrefix `
                                                           -SourcePortRange $rule.SourcePortRange `
                                                           -DestinationAddressPrefix $rule.DestinationAddressPrefix `
                                                           -DestinationPortRange $rule.DestinationPortRange    
   
        # Build an initial list of network security groups
        if ($networkSecurityGroups.Keys -notcontains $rule.networkSecurityGroupName)
        {
            $ruleList = New-Object System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSecurityRule]
            $ruleList.Add($ruleConfig)
            $networkSecurityGroups.Add(($rule.networkSecurityGroupName), ($ruleList))
        }
    
        else
        {
            $expression = '$networkSecurityGroups.' + "'" + $rule.networkSecurityGroupName +"'" + '.Add($ruleConfig)'
            Invoke-Expression $expression
        }      
    }

    Write-Progress -Activity "Getting network security rules.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any network security rules to create.'
}

$nsgs =  Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                      -SheetName "Network Security Groups" `
                      -closeExcel

if($nsgs -ne $null)
{
    # Create network security group objects
    foreach ($nsg in $nsgs)
    {
        $nsgName = $nsg.name
        Write-Progress -Activity "Creating network security groups.." `
                       -Status "Working on $nsgName" `
                       -PercentComplete ((($nsgs.IndexOf($nsg)) / $nsgs.Count) * 100)
   
        $expression = '$networkSecurityGroups.' + "'" + "$nsgName" + "'"
        New-AzureRmNetworkSecurityGroup -ResourceGroupName $nsg.resourceGroupName`
                                        -Location $nsg.location `
                                        -Name $nsg.name `
                                        -SecurityRules (Invoke-Expression $expression) `
                                        -WarningAction Ignore | Out-Null
    }

    Write-Progress -Activity "Creating network security groups.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any network security groups to create.'
}

[array]$publicIPs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                 -SheetName "Public IPs" `
                                 -closeExcel

if($publicIPs -ne $null)
{

    Write-Host ''
    Write-Host 'Will now create public IP(s).'


    # Create the puclic IPs
    foreach ($ip in $publicIPs)
    {
        $ipName = $ip.name
        Write-Progress -Activity "Creating public IPs.." `
                       -Status "Working on $ipName" `
                       -PercentComplete ((($publicIPs.IndexOf($ip)) / $publicIPs.Count) * 100)
    
        New-AzureRmPublicIpAddress -Name $ip.name `
                                   -ResourceGroupName $ip.resourceGroupName `
                                   -Location $ip.location `
                                   -AllocationMethod $ip.allocationMethod `
                                   -WarningAction Ignore | Out-Null
}

    Write-Progress -Activity "Creating public IPs.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed

}

else
{
    Write-Host ''
    Write-Host 'Did not find any public IPs to create.'    
}

[array]$vngIpConfigs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                    -SheetName "VNG IP Config" `
                                    -closeExcel

[array]$vngs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                            -SheetName "Virtual Network Gateways" `
                            -closeExcel

if($vngIpConfigs -ne $null -and $vngs -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the virtual network gateway(s).'

    # Configure the virtual network gateways
    foreach ($vngIpConfig in $vngIpConfigs)
    {
        $vngIpConfigVnetName = $vngIpConfig.VNET
        
        $gatewayVnet = Get-AzureRmVirtualNetwork -Name $vngIpConfig.VNET `
                                                 -ResourceGroupName ($vnets | Where-Object {$_.vnetName -like "$vngIpConfigVnetName"} | Select-Object -ExpandProperty resourceGroupName)

        $gatewayVnetName = $gatewayVnet.Name
    
        $subnet = Get-AzureRmVirtualNetworkSubnetConfig -Name $vngIpConfig.subnet `
                                                        -VirtualNetwork $gatewayVnet
    
        $gwipconfig = New-AzureRmVirtualNetworkGatewayIpConfig -Name $vngIpConfig.name `
                                                               -SubnetId $subnet.Id `
                                                               -PublicIpAddressId (Get-AzureRmPublicIpAddress -Name $vngIpConfig.publicIP `
                                                                                                          -ResourceGroupName ($publicIPs | Where-Object {$_.name -like $vngIpConfig.publicIP} | Select-Object -ExpandProperty resourceGroupName)).Id

        $vng = $vngs | Where-Object {$_.name -like $vngIpConfig.name} 
        $vngName = $vng.Name
        Write-Progress -Activity "Creating virtual network gateways.." `
                       -Status "Working on $vngName" `
                       -PercentComplete ((($vngIpConfigs.IndexOf($vngIpConfig)) / $vngIpConfigs.Count) * 100)
    
        New-AzureRmVirtualNetworkGateway -Name $vng.name `
                                         -ResourceGroupName $vng.resourceGroupName `
                                         -Location $vng.location `
                                         -IpConfigurations $gwipconfig `
                                         -GatewayType $vng.gatewayType `
                                         -VpnType $vng.vpnType `
                                         -GatewaySku $vng.gatewaySku `
                                         -WarningAction Ignore | Out-Null
    }

    Write-Progress -Activity "Creating virtual network gateways.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any virtual network gateways to create.'
}

if($vnets.Count -gt 1 -and $gatewayVnetName -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create network peering between VNETs.'
    Write-Host ''
    Write-Host 'Note that a hub/spoke archicture is configured with the VNET containing'
    Write-Host 'the virtual network gateway acting as the hub.'

    # Configure peering between all VNETs - VNET with gateway to on-premises 
    # address space acts as hub. Note that VNETs without a gatway are not able
    # to use the gateway to directly access on-premises address space
    [array]$vnets = Get-AzureRmVirtualNetwork
    $vnets = $vnets | Where-Object {$_.Name -notlike $gatewayVnet.Name}

    # Create the hub and spokes
    foreach ($vnet in $vnets)
    {
        $vnetName = $vnet.Name
        Write-Progress -Activity "Creating peering connections.." `
                       -Status "Working on $gatewayVnetName-$vnetName connection" `
                       -PercentComplete ((($vnets.IndexOf($vnet)) / $vnets.Count) * 100)
    
        Add-AzureRmVirtualNetworkPeering -Name "azrVNP$gatewayVnetName-$vnetName" `
                                         -VirtualNetwork $gatewayVnet `
                                         -RemoteVirtualNetworkId $vnet.Id `
                                         -WarningAction Ignore | Out-Null

        Add-AzureRmVirtualNetworkPeering -Name "azrVNP$vnetName-$gatewayVnetName" `
                                         -VirtualNetwork $vnet `
                                         -RemoteVirtualNetworkId $gatewayVnet.Id `
                                         -WarningAction Ignore | Out-Null
    }

    Write-Progress -Activity "Creating peering connections.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find enough VNETs or a VNET with a virtual network gateway'
    Write-Host 'to create a hub/spoke archicture with virtual network gateway acting'
    Write-Host 'as the hub. Note that peering can still be configured to connect VNETs,'
    Write-Host 'but it must be configured manually.' 
}

$lngs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                     -SheetName "Local Network Gateways" `
                     -closeExcel

$vngConnections = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                               -SheetName "VNG Connections" `
                               -closeExcel

if($lngs -ne $null -and $vngConnections -ne $null)
{
    Write-Host ''
    Write-Host 'Will now establish site-to-site VPN connections.'

    # Configre the virtual network gateways
    foreach ($lng in $lngs)
    {
        $lngName - $lng.name
        $lngGatewayIpAddress = $lng.gatewayIpAddress
        $lngAddressPrefix = $lng.addressPrefix
        Write-Progress -Activity "Creating site-to-site connections.." `
                       -Status "Working on connection to $lngGatewayIpAddress / $lngAddressPrefix" `
                       -PercentComplete ((($lngs.IndexOf($lng)) / $lngs.Count) * 100) 
        $localNetworkGateway = New-AzureRmLocalNetworkGateway -Name $lng.name `
                                                              -ResourceGroupName $lng.resourceGroupName `
                                                              -Location $lng.location `
                                                              -GatewayIpAddress $lng.gatewayIpAddress `
                                                              -AddressPrefix $lng.addressPrefix | Out-Null

        $vngConnection = $vngConnections | Where-Object {$_.localNetworkGateway2 -like "$lngName"}
        $virtualNetworkGateway = $vngs | Where-Object {$_.name -like $lng.azureVng}
        $viritulNetworkGateway = Get-AzureRmVirtualNetworkGateway -Name $virtualNetworkGateway.name `
                                                                  -ResourceGroupName $virtualNetworkGateway.resourceGroupName `
        $secureString = ConvertTo-SecureString –String $vngConnection.sharedSecret `
                                               –AsPlainText `
                                               -Force                                                          
    
        New-AzureRmVirtualNetworkGatewayConnection -Name $vngConnection.name `
                                                   -ResourceGroupName $vngConnection.resourceGroupName `
                                                   -Location $vngConnection.location `
                                                   -VirtualNetworkGateway1 $virtualNetworkGateway `
                                                   -LocalNetworkGateway2 $localNetworkGateway `
                                                   -ConnectionType $vngConnection.connectionType `
                                                   -RoutingWeight $vngConnection.routingWeight `
                                                   -SharedKey $secureString | Out-Null
}

    Write-Progress -Activity "Creating site-to-site connections.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
}

else
{
    Write-Host ''
    write-Host 'Did not find virtual gateway connections to establish site-to-site VPNs.'
}

Write-Host ''
Write-Host 'The virtual network and sub-components have been created.'
Write-Host 'The next step is to create any needed VMs in the environment.'
Write-Host 'After VMs are created, NSGs can be attached to NICs and subnets.'
