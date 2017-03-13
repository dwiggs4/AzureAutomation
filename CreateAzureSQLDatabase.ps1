# Include Azure Resource Manager Function Library
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
Write-Host 'Please select a subscription for the new Azure SQL Database.'
$objSubscription = SelectAzureRmSubscription

# Collect required input
Write-Host ''
Write-Host 'Please select a region for the new Azure SQL Database.'
$arrRegions = 'East US','West US'
$strLocation = StringPicker($arrRegions)

# Get/create resource group
Write-Host ''
Write-Host 'Please select a resource group for the new Azure SQL Database.'
Write-Host 'The following resource groups exist in this subscription:'
$objResourceGroup = ObjectPicker(Get-AzureRmResourceGroup)
if($objResourceGroup -eq $null)
{
    CreateResourceGroup
}

# Get Azure SQL server name, version, and administrative credential
do
{
    Write-Host ''
    $strUseExistingSQLServer = Read-Host -Prompt 'Use and existing Azure SQL server? (Y/N)'
}
until ($strUseExistingSQLServer.ToLower() -eq 'y' -or $strUseExistingSQLServer.ToLower() -eq 'n')

if ($strUseExistingSQLServer.ToLower() -eq 'y')
{
    $objSqlDbServer = ObjectPicker(Get-AzureRmSqlServers)
    $strServerName = $objSqlDbServer.ServerName
}

if ($objSqlDbServer -eq $null -or $strUseExistingSQLServer.ToLower() -eq 'n')
{
    $strServerName = GetAzureSQLServerName
    $arrServerVersions = '11.0','12.0'
    Write-Host 'The alignment of SQL versions to default compatibility levels are as follows:'
    Write-Host '' 
    Write-Host 'SQL Server 2008 -> Azure SQL Database 11.0.'
    Write-Host 'SQL Server 2012 -> Azure SQL Database 11.0.'
    Write-Host 'SQL Server 2014 -> Azure SQL Database 12.0.'
    Write-Host 'SQL Server 2016 -> Azure SQL Database 12.0.'
    Write-Host ''
    Write-Host 'Please select aversion for the new Azure SQL Server.'
    $strServerVersion = StringPicker($arrServerVersions)
    $objSqlAdministratorCredential = Get-Credential –Message "Please enter the desired local administrator password." `
                                                    -UserName ("adm."+($strServerName.ToUpper()))

    # Create the Azure SQL server
    Write-Host ''
    Write-Host 'Creating the Azure SQL Server...'
    try
    {
        $objSqlDbServer = New-AzureRmSqlServer -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                               -ServerName $strServerName `
                                               -Location $strLocation `
                                               -ServerVersion $strServerVersion `
                                               -SqlAdministratorCredentials $objSqlAdministratorCredential | Out-Null
        
        Write-Host 'Successfully created the Azure SQL Server...' -ForegroundColor Green
    }
    catch 
    {
        Write-Host 'Unable to create the Azure SQL Server...' -ForegroundColor Red
        exit
    }

    # Determine whether a recovery services vault should be used for backing up databases
    do
    {
        Write-Host ''
        $strUseRSV = Read-Host -Prompt 'Backup databases in the new Azure SQL Server to a Recovery Services Vault? (Y/N)'
    }
    until ($strUseRSV.ToLower() -eq 'y' -or $strUseRSV.ToLower() -eq 'n')

    if ($strUseRSV.ToLower() -eq 'y')
    {
        do
        {
        $strUseExistingRSV = Read-Host -Prompt 'Use and existing Recovery Services Vault? (Y/N)'
        }
        until ($strUseExistingRSV.ToLower() -eq 'y' -or $strUseExistingRSV.ToLower() -eq 'n')

        if ($strUseExistingRSV.ToLower() -eq 'y')
        {
            Write-Host 'Note that the vault must be located in the' -NoNewline; Write-Host ' same region' -ForegroundColor Yellow -NoNewline; Write-Host ' as the Azure SQL logical server,' 
            Write-Host 'and must use the' -NoNewline; Write-Host ' same resource group' -ForegroundColor Yellow -NoNewline; ' as the logical server.'
            Write-Host ''
            Write-Host 'Please select an existing Recovery Services Vault for the new Azure SQL Server.'
            $objRecoveryServicesVault = ObjectPicker(Get-AzureRmRecoveryServicesVault)
        }
        if ($objRecoveryServicesVault -eq $null -or $strUseExistingRSV -eq 'n')
        {
            $resourceProviders = Get-AzureRmResourceProvider
            if (($resourceProviders | Select-Object -ExpandProperty ProviderNamespace) -notcontains 'Microsoft.RecoveryServices')
            {
                Register-AzureRmResourceProvider -ProviderNamespace Microsoft.RecoveryServices
            }

            $strVaultName = GetRecoveryServicesVaultName
            Write-Host 'Creating RSV...'
            try
            {
                $objRecoveryServicesVault = New-AzureRmRecoveryServicesVault -Name $strVaultName `
                                                                                -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                                                -Location $strLocation

                Write-Host 'Successfully created the RSV...' -ForegroundColor Green
            }
            catch 
            {
                Write-Host "Unable to create the RSV..." -ForegroundColor Red
                exit
            }
            
            # Specify storage redundancy for the vault
            $arrBackupStorageRedundancy = 'GeoRedundant','LocallyRedundant'
            Write-Host ''
            Write-Host 'Please select the type of storage redundancy.'
            $strBackupStorageRedundancy = StringPicker($arrBackupStorageRedundancy)
            Set-AzureRmRecoveryServicesBackupProperties -BackupStorageRedundancy $strBackupStorageRedundancy `
                                                        -Vault $objRecoveryServicesVault
            Set-AzureRmSqlServerBackupLongTermRetentionVault -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                                -ServerName $strServerName `
                                                                –ResourceId $objRecoveryServicesVault.Id

            # Retrieve the default in-memory policy object for AzureSQLServer workload and set the retention period
            $objRetentionPolicy = Get-AzureRmRecoveryServicesBackupRetentionPolicyObject -WorkloadType AzureSQLDatabase
            # Decide on how long to keep the backups
            # The presentation of how long to keep backups could be more elegant 
            $arrRetentionDurationType = 'Months' ,`
                                        'Years'
            Write-Host ''
            Write-Host 'Please select the unit for the retention policy.'
            $strRetentionDurationType = StringPicker($arrRetentionDurationType)
            $arrRetentionCount = 1..12
            Write-Host ''
            Write-Host 'Please select the length for the retention policy.'
            $strRetentionCount = StringPicker($arrRetentionCount)
            $objRetentionPolicy.RetentionDurationType = $retentionDurationType
            $objRetentionPolicy.RetentionCount = $strRetentionCount

            # Register the policy for use with SQL databases
            Set-AzureRMRecoveryServicesVaultContext -Vault $objRecoveryServicesVault
            $policyName = Read-Host -Prompt 'Please provide a name for the backup policy'
            Write-Host 'Creating the backup policy...'
            try
            {
                $policy = New-AzureRmRecoveryServicesBackupProtectionPolicy -Name $policyName `
                                                                            –WorkloadType AzureSQLDatabase `
                                                                            -RetentionPolicy $objRetentionPolicy

                Write-Host 'Successfully created the backup policy...' -ForegroundColor Green
            }
            catch 
            {
                Write-Host "Unable to create the backup policy..." -ForegroundColor Red
                exit
            }
        }
    }
}

<#
# ToDo - build this out more
# Create an Azure SQL Server firewall rule for the production ETL server
# Note that this firewall rule is a server rule - the production ETL server
# will have access to all the databases on the Azure SQL server
[string]$ip = '10.1.2.5'
[string]$firewallRuleName = 'AZBPRDETL001-AZBASSPRD001'
[string]$firewallStartIp = $ip
[string]$firewallEndIp = $ip

$fireWallRule = New-AzureRmSqlServerFirewallRule -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                 -ServerName $strServerName `
                                                 -FirewallRuleName $firewallRuleName `
                                                 -StartIpAddress $firewallStartIp `
                                                 -EndIpAddress $firewallEndIp
#>

# Get database name
$strDatabaseName = GetAzureSQLDatabaseName

# Build pricing details object to make decision of pricing tier and requested service level objective easier
{
$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Basic'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'Basic'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '5'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '2'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.0067/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Standard'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'S0'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '10'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.0202/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Standard'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'S1'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '20'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.0403/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Standard'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'S2'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '50'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.1008/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Standard'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'S3'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '100'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.2016/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P1'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '125'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.6250/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P2'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$1.250/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P4'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$2.5000/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P6'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '1,000'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$5.000/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P11'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '1,700'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '1,024'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$9.4100/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'P15'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '4,000'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '1,024'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$21.5100/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium RS'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'PRS1'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '125'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.1563/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium RS'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'PRS2'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.3125/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium RS'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'PRS4'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '500'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$0.6250/hr'
[array]$arrPricingDetails += $objPricingDetail

$objPricingDetail = New-Object -TypeName PSObject
$objPricingDetail | Add-Member -Name 'PricingTier' -MemberType Noteproperty -Value 'Premium RS'
$objPricingDetail | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value 'PRS6'
$objPricingDetail | Add-Member -Name 'DTUs' -MemberType Noteproperty -Value '1,000'
$objPricingDetail | Add-Member -Name 'Storage (GB)' -MemberType Noteproperty -Value '250'
$objPricingDetail | Add-Member -Name 'Price' -MemberType Noteproperty -Value '$1.2500/hr'
[array]$arrPricingDetails += $objPricingDetail
 
$arrPricingDetails | FT -AutoSize
}

# Get information about database object
function GetAzureSQLDatabaseConfiguration
{
    # Build custom object to store database details to make using a for loop easier
    $objDatabase = New-Object -TypeName PSObject
    $objDatabase | Add-Member -Name 'DatabaseName' -MemberType Noteproperty -Value $strDatabaseName

    $arrEdition = 'Basic','Premium', 'PremiumRS','Standard'
    
    Write-Host 'Please select an edition for the Azure SQL Database'
    $strEdition = StringPicker($arrEdition)
    $objDatabase | Add-Member -Name 'Edition' -MemberType Noteproperty -Value $strEdition
    
    if ($arrEdition -eq 'Basic')
    {
        # Cast variable to array to ensure that StringPicker works correctly
        [array]$arrRequestedServiceObjectiveName = 'Basic'
    }
    elseif ($arrEdition -eq 'Premium')
    {
        $arrRequestedServiceObjectiveName = 'P1','P2','P4','P6','P11'
    }
    elseif ($arrEdition -eq 'PremiumRS')
    {
        $arrRequestedServiceObjectiveName = 'PRS1','PRS2','PRS4','PRS6'
    }
    else
    {
        $arrRequestedServiceObjectiveName = 'S0','S1','S2','S3'
    }
    
    Write-Host ''
    Write-Host 'Please select requested service objective for the Azure SQL Database'
    $strRequestedServiceObjectiveName = StringPicker($arrRequestedServiceObjectiveName)
    $objDatabase | Add-Member -Name 'RequestedServiceObjective' -MemberType Noteproperty -Value $strRequestedServiceObjectiveName

    return $objDatabase
}

$databaseToAdd = GetAzureSQLDatabaseConfiguration

[array]$arrDatabases += $databaseToAdd

do
{
    Write-Host ''
    $addDatabase = Read-Host -Prompt 'Would you like to add another database? (Y/N)'

    if ($addDatabase.ToLower() -eq 'y')
    {
        $strDatabaseName = Read-Host -Prompt 'Please provide the object name for the new Azure SQL Database'
        $databaseToAdd = GetAzureSQLDatabaseConfiguration
        [array]$arrDatabases += $databaseToAdd
    }
}
until ($addDatabase.ToLower() -eq 'n')

# Create the Azure Sql Databases, enable TDE, and set retention policy if applicable
Write-Host ''
Write-Host 'Creating databases...'
foreach ($database in $arrDatabases)
{
    $databaseName = $database.DatabaseName
    Write-Progress -Activity "Creating Azure SQL Databases.." `
                   -Status "Working on databse: $databaseName" `
                   -PercentComplete ((($arrDatabases.IndexOf($database)) / $arrDatabases.Count) * 100)

    try
    {
        New-AzureRmSqlDatabase -ResourceGroupName $objResourceGroup.ResourceGroupName `
                               -ServerName $strServerName `
                               -DatabaseName $database.DatabaseName `
                               -Edition $database.Edition `
                               -RequestedServiceObjectiveName $database.RequestedServiceObjective
        
        Write-Host "Successfully created database $databaseName" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Unable to create database $databaseName" -ForegroundColor Red
    }
    
    try
    {
        Write-Host ''
        Write-Host "Enabling Transparent Data Encryption (TDE) for database $databaseName..."

        Set-AzureRMSqlDatabaseTransparentDataEncryption -ServerName $strServerName `
                                                        -ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                        -DatabaseName $database.DatabaseName `
                                                        -State "Enabled" | Out-Null

        Write-Host "Successfully enabled Transparent Data Encryption (TDE) for database $databaseName..." -ForegroundColor Green
    }
    catch
    {
        Write-Host "Unable to enable Transparent Data Encryption (TDE) for database $databaseName" -ForegroundColor Red
    }

    if ($strUseRSV -eq 'y')
    {
        try
        {
            Write-Host ''
            Write-Host "Setting the long term retention policy for database $databaseName"

            Set-AzureRmSqlDatabaseBackupLongTermRetentionPolicy –ResourceGroupName $objResourceGroup.ResourceGroupName `
                                                                –ServerName $strServerName `
                                                                -DatabaseName $database.DatabaseName `
                                                                -State "Enabled" `
                                                                -ResourceId $policy.Id | Out-Null 

            Write-Host "Successfully set the long term retention policy for database $databaseName"
        }
        catch
        {
            Write-Host "Unable to set the long term retention policy for database $databaseName" -ForegroundColor Red
        }
    }
    
}
Write-Progress -Activity "Creating Azure SQL Databases.." `
               -Status "Working on databse: $databaseName" `
               -PercentComplete 100 `
               -Completed