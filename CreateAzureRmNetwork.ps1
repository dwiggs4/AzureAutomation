function Open-ExcelApplication 
{
    param
    (
        [switch] $Visible,
        [switch] $HideAlerts
    ) 
    
    $app = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
    $app.Visible  = $Visible
    $app.DisplayAlerts = -not $HideAlerts
    return $app
}

function Get-ExcelWorkBook 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        [switch] $Visible,
        [switch] $readonly
    )

    [Microsoft.Office.Interop.Excel.ApplicationClass]$app = $null 
    if($inputObject -is [Microsoft.Office.Interop.Excel.ApplicationClass]) 
    {
        $app = $inputObject
        $WorkBook = $app.ActiveWorkbook
    } 
    
    else 
    {
        $app = Open-ExcelApplication -Visible:$Visible  
        try 
        {
            if($inputObject.Contains("\\") -or $inputObject.Contains("//")) 
            {
                $WorkBook = $app.Workbooks.Open($inputObject,$true,[System.Boolean]$readonly)
            } 
            
            else 
            {
                $WorkBook = $app.Workbooks.Open((Resolve-path $inputObject),$true,[System.Boolean]$readonly)
            }
        } 
        
        catch 
        {
            $WorkBook = $app.Workbooks.Open((Resolve-path $inputObject),$true,[System.Boolean]$readonly)
        }
    } 

    $app.CalculateFullRebuild() 
    return $WorkBook
}

function Get-ExcelWorkSheet 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        $SheetName,
        [switch] $Visible,
        [switch] $readonly
    )
  
    if($inputObject -is [Microsoft.Office.Interop.Excel.Workbook]) 
    {
        $WorkBook = $inputObject
    } 
    
    else 
    {
        $WorkBook = Get-ExcelWorkBook $inputObject -Visible:$Visible `
                                                   -readonly:$readonly
    }
    
    if (($SheetName -eq $null) -or $SheetName -eq 0) 
    {
        $WorkBook.ActiveSheet
    } 
    
    else 
    {
        $WorkBook.WorkSheets.item($SheetName)
    } 
}

function Import-Row 
{
    param
    (
        $Row,[hashtable] $Headers =@{},
        $ColumnStart = 1,
        $ColumnCount = $Row.Value2.Count
    )
    
    $output = @{}
    for ($index=$ColumnStart;$index -le $ColumnCount;$index ++)
    {
        if($Headers.Count -eq 0)
        {
            $Key = $Index
        } 
        
        Else 
        {
            $Key = $Headers[$index]
        }
        
        $output.Add($Key,$row.Cells.Item(1,$index).Text)
    }
    return $output
}

function Release-Ref ($ref) 
{
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)  | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null
}

function Close-ExcelApplication 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject
    )
    
    if ($inputObject -is [Microsoft.Office.Interop.Excel.ApplicationClass]) 
    {
        $app = $inputObject  
    } 
    else 
    {
    $app = $inputObject.Application
    Release-Ref $inputObject
    }

    $app.ActiveWorkBook.Close($false) | Out-Null
    $app.Quit() | Out-Null
    Release-Ref $app
}

Function Import-Excel 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        [Object] $SheetName,
        [switch] $Visible,
        [switch] $readonly,
        [int] $startOnLineNumber =1,
        [switch] $closeExcel,
        [switch] $asHashTable,
        [hashtable] $FieldNames =@{}
    )
    
    #Check what the input is. 
    if ($inputObject -is [Microsoft.Office.Interop.Excel.range]) 
    { 
        $range = $inputObject
    } 
    elseif ($inputObject -isnot [Microsoft.Office.Interop.Excel.Worksheet]) 
    { 
        $WorkSheet = Get-ExcelWorkSheet $inputObject -SheetName $SheetName `
                                                     -Visible:$Visible `
                                                     -readonly:$readonly  
        $range = $WorkSheet.UsedRange
    } 
    else 
    {
        $WorkSheet = $inputObject
        $range = $WorkSheet.UsedRange
    }
    
    # populate the Header 
    if ($FieldNames.Count -eq 0) 
    {
        $FieldNames = Import-Row $range.Rows.Item($startOnLineNumber++)              
    }

    for ($RowIndex=$startOnLineNumber;$RowIndex -le $range.Rows.Count;$RowIndex++) 
    {
        $output = Import-Row $range.Rows.Item($RowIndex) -Headers $FieldNames
    
        if ($asHashtAble) 
        {
            Write-Output $output
        } 
        else 
        {
            New-Object PSObject -property $output
        }
    }  

    # If we opened Excel, we should close Excel.
    if ($closeExcel) 
    {   
        $WorkSheet.Activate() | Out-Null
        Close-ExcelApplication $WorkSheet
    } 
}

# Script requires WMPAzureRMFunctionalLibrary.ps1
if(Test-Path .\WMPAzureRMFunctionLibrary.ps1)
{
    # Include Azure RM Function Library
    . .\WMPAzureRMFunctionLibrary.ps1
}
else
{
    Write-Host 'Cannot find helper library WMPAzureRMFunctionLibrary.ps1.'
    Write-Host 'Please ensure the working directory includes the helper library.'
    Write-Host 'Ending Script.'
    exit
}

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

Write-Host 'Will now create the resource group(s).'

$resourceGroups = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                               -SheetName "Resource Groups" `
                               -closeExcel

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

Write-Host ''
Write-Host 'Will now create the storage account(s).'

$storageAccounts = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                -SheetName "Storage Accounts" `
                                -closeExcel

foreach ($storageAccount in $storageAccounts)
{
    $storageAccountName = $storageAccount.storageAccountName
    Write-Progress -Activity "Creating storage accounts.." `
                   -Status "Working on $storageAccountName" `
                   -PercentComplete ((($storageAccounts.IndexOf($storageAccount)) / $storageAccounts.Count) * 100)
    
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

Write-Host ''
Write-Host 'Will now create the subnet(s).'

$subnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                        -SheetName "Subnets" `
                        -closeExcel
  
# Create hashtable that groups virtual networks and associated subnets
$virtualNetworks = @{}
foreach ($subnet in $subnets)
{
    $vnetName = $subnet.VNET
    $subnetName = $subnet.subnetName
    
    if($virtualNetworks.Keys -notcontains $subnet.VNET)
    {
        $subnetList = New-Object System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSubnet]
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

Write-Host ''
Write-Host 'Will now create the VNET(s).'

$vnets = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                      -SheetName "VNETs" `
                      -closeExcel

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

# Initialize empty hash table
Write-Host ''
Write-Host 'Will now create the network security group(s).'
$networkSecurityGroups = @{}
$networkSecurityRules = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                     -SheetName "Network Security Rules" `
                                     -closeExcel
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

# Create network security group objects
$nsgs =  Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                      -SheetName "Network Security Groups" `
                      -closeExcel

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

Write-Host ''
Write-Host 'Will now create public IP(s).'

[array]$publicIPs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                 -SheetName "Public IPs" `
                                 -closeExcel

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

Write-Host ''
Write-Host 'Will now create the virtual network gateway(s).'

# Configure the virtual network gateways
[array]$vngIpConfigs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                                    -SheetName "VNG IP Config" `
                                    -closeExcel

[array]$vngs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                            -SheetName "Virtual Network Gateways" `
                            -closeExcel

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

Write-Host ''
Write-Host 'Will now establish site-to-site VPN connections.'

# Configre the virtual network gateways
$lngs = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                     -SheetName "Local Network Gateways" `
                     -closeExcel

$vngConnections = Import-Excel -inputObject .\AzureRmNetwork.xlsx `
                               -SheetName "VNG Connections" `
                               -closeExcel

foreach ($lng in $lngs)
{
    $local = New-AzureRmLocalNetworkGateway -Name $lng.name `
                                            -ResourceGroupName $lng.resourceGroupName `
                                            -Location $lng.location `
                                            -GatewayIpAddress $lng.gatewayIpAddress `
                                            -AddressPrefix $lng.addressPrefix | Out-Null

    New-AzureRmVirtualNetworkGatewayConnection -Name azbGCNazbVNTBIEnvironment-azbLNGDENFW `
                                               -ResourceGroupName azbRSGNetwork `
                                               -Location 'West US' `
                                               -VirtualNetworkGateway1 $gateway1 `
                                               -LocalNetworkGateway2 $local `
                                               -ConnectionType IPsec `
                                               -RoutingWeight 10 `
                                               -SharedKey 'abc123' | Out-Null


}

Write-Host 'The virtual network, subnets, and network security groups and rules have been created.'
Write-Host 'The next step is to create any needed VMs in the environment.'
Write-Host 'After VMs are created, NSGs can be attached to NICs and subnets.'
