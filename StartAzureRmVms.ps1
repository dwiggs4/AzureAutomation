<# 
Script to start all VMs in a given resource group.
The workflow relies on scheduling to be run at a specific time.
Note that the credential "AzureCredential" linked to the automation account uses the
svc.azureautomation@bridgehealthmedical.com credential to perform tasks.
The svc.azureautomation@bridgehealthmedical credential has been given the role of 
VM contributer at the subscription level.
#>

$credential  = Get-AutomationPSCredential -Name "AzureCredential"
Add-AzureRmAccount -Credential $credential | Out-Null
Select-AzureRmSubscription -SubscriptionName 'BI Environment' | Out-Null
$resourceGroupName = 'azbRSGDEVCompute'
$VMs = Get-AzureRmVM -ResourceGroupName $resourceGroupName

$output = 'All the VMs in the ' + $resourceGroupName + ' resource group in the VM deallocated state will be started'
Write-Output $output

foreach ($VM in $VMs)
{
    $powerStateStatus = (Get-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -Status).Statuses[1]
    if($powerStateStatus.DisplayStatus -like 'VM deallocated')
    {
        $output = 'INFO: Starting VM ' + $VM.Name + '.'
        Write-Output $output
        Start-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name | Out-Null
        $powerStateStatus = (Get-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -Status).Statuses[1]

        if($powerStateStatus.DisplayStatus -like 'VM running')
        {
            $output = 'INFO: ' + $VM.Name + ' started.'
            Write-Output $output
        }

        else
        {
            $output = 'ERROR: ' + $VM.Name + ' failed to start.'
            Write-Output $output         
        }         
    }

    else
    {
        $output = 'INFO: ' + $VM.Name + ' is already started.'
        Write-Output $output         
    }
}