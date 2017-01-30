# Include Excel Function Library
."$PSScriptRoot\ExcelFunctionalLibrary.ps1"

# Script requires AzureRmNetwork.xlsx
if(-not (Test-Path .\AzureRmNetwork.xlsx))
{
    Write-Host 'Cannot find helper networkSecurityRules.csv.'
    Write-Host 'Please ensure the working directory includes the CSV.'
    Write-Host 'Ending Script.'
    exit
}

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

$resourceGroups = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                               -SheetName "Resource Groups" `
                               -closeExcel

if($resourceGroups -ne $null)
{
    Write-Host 'Will now create the resource group(s).'


    foreach ($resourceGroup in $resourceGroups)
    {
        $resourceGroupName = $resourceGroup.resourceGroupName
        Write-Progress -Activity "Creating resource groups.." `
                       -Status "Working on $resourceGroupName" `
                       -PercentComplete ((($resourceGroups.IndexOf($resourceGroup)) / $resourceGroups.Count) * 100)
    
        New-AzureRmResourceGroup -Name $resourceGroupName `
                                 -Location $resourceGroup.location | Out-Null
}

Write-Progress -Activity "Creating resource groups.." `
               -Status "Done" `
               -PercentComplete 100 `
               -Completed
}

else
{
    Write-Host ''
    Write-Host 'Did not find any resource groups to create.'
}

$storageAccounts = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                -SheetName "Storage Accounts" `
                                -closeExcel
if($storageAccounts -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the storage account(s).'

    foreach ($storageAccount in $storageAccounts)
    {
        $storageAccountName = $storageAccount.storageAccountName
        Write-Progress -Activity "Creating storage accounts.." `
                       -Status "Working on $storageAccountName" `
                       -PercentComplete ((($storageAccounts.IndexOf($storageAccount)) / $storageAccounts.Count) * 100)
    
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

$subnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                        -SheetName "Subnets" `
                        -closeExcel

if($subnets -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the subnet(s).'
  
    # Create hashtable that groups virtual networks and associated subnets
    $virtualNetworks = @{}
    foreach ($subnet in $subnets)
    {
        $vnetName = $subnet.VNET
        $subnetName = $subnet.subnetName
    
        if($virtualNetworks.Keys -notcontains $subnet.VNET)
        {
            $subnetList = New-Object 'System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSubnet]'
            $virtualNetworks.Add("$vnetName", $subnetList)
            Write-Progress -Activity "Creating network subnets.." `
                           -Status "Working on $subnetName" `
                           -PercentComplete ((($subnets.IndexOf($subnet)) / $subnets.Count) * 100)
            $subnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name $subnet.subnetName `
                                                                  -AddressPrefix $subnet.CIDR
            $expression = '$virtualNetworks.' + "$vnetName" + '.Add($subnetConfig)'
            Invoke-Expression $expression
        }

        else
        {
            Write-Progress -Activity "Creating network subnets.." `
                           -Status "Working on $subnetName" `
                           -PercentComplete ((($subnets.IndexOf($subnet)) / $subnets.Count) * 100)
            $subnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name $subnet.subnetName `
                                                                  -AddressPrefix $subnet.CIDR
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
    Write-Host ''
    Write-Host 'Did not find any subnets to create.'
}

$vnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                      -SheetName "VNETs" `
                      -closeExcel

if($vnets -ne $null)
{
    Write-Host ''
    Write-Host 'Will now create the VNET(s).'

    # Create virtual networks
    foreach ($vnet in $vnets)
    {
        $vnetName = $vnet.vnetName    
        Write-Progress -Activity "Creating virtual networks.." `
                       -Status "Working on $vnetName" `
                       -PercentComplete ((($vnets.IndexOf($vnet)) / $vnets.Count) * 100)

        $expression = '$virtualNetworks.' + "$vnetName"
        $subnetList = (Invoke-Expression $expression)

        New-AzureRmVirtualNetwork -Name $vnet.vnetName `
                                  -ResourceGroupName $vnet.resourceGroupName `
                                  -Location $vnet.location `
                                  -AddressPrefix $vnet.CIDR `
                                  -Subnet $subnetList `
                                  -WarningAction Ignore | Out-Null    
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

$networkSecurityRules = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
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

if($vnets.Count -gt 1)
{
    Write-Host ''
    Write-Host 'Will now create network peering between VNETs.'

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
    Write-Host 'Did not find enough VNETs to create a peering mesh.'
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

Write-Host 'The virtual network and sub-components have been created.'
Write-Host 'The next step is to create any needed VMs in the environment.'
Write-Host 'After VMs are created, NSGs can be attached to NICs and subnets.'
