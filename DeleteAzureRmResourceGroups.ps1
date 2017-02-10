# Include Azure RM Function Library
."$PSScriptRoot\AzureRmFunctionalLibrary.ps1"

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

$resourceGroups = Get-AzureRmResourceGroup

Write-Host 'The following resource groups exist:'

$resourceGroups | FT ResourceGroupName

$response = Read-Host 'Would you like to delete all resource groups? (yes/no)' 
$validResponse = 0 
while($validResponse -eq 0){
    if($response -like 'yes') {
        Write-Host ''
        Write-Host 'Will delete all resource groups.'
        foreach ($group in $resourceGroups)
        {
            $groupName = $group.ResourceGroupName
            Write-Progress -Activity "Deleting resource groups.." `
                           -Status "Working on $groupName" `
                           -PercentComplete ((($resourceGroups.IndexOf($group)) / $resourceGroups.Count) * 100)
            
            try
            {
                 Remove-AzureRmResourceGroup -Id $group.ResourceId -Force -ErrorAction Stop | Out-Null
                 Write-Debug "Removed resource group $groupName" 
            }
            catch
            {
                $errorMessage = $_.Exception.Message
                Write-Host "Error removing $groupName"
                Write-Debug "$errorMessage"
            }
        }
        $validResponse = 1        
}
            elseif($response -like 'no'){
            Write-Host ''
            Write-Host 'Will not delete resource groups.'            
            $validResponse = 1        
            }

            else{
            Write-Host "Please specify 'yes' or 'no'."
            $response = Read-Host 'Would you like to delete all resource groups? (yes/no)'
            }
}



# ToDo
# Build out logic for RSV teardown/deletion
