<# 
Script to start all VMs in a given resource group in parallel
The workflow relies on scheduling to be run at a specific time.
The default name of "AzureRunAsConnection" is used as a RunAs account was created for the Azure Automation Account
#>

workflow azrARBstart-<resourceGroup>Vms
{   
    $connectionName = "AzureRunAsConnection"
    try
    {
        # Get the connection "AzureRunAsConnection "
        $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         

        "Logging in to Azure..."
        Add-AzureRmAccount `
            -ServicePrincipal `
            -TenantId $servicePrincipalConnection.TenantId `
            -ApplicationId $servicePrincipalConnection.ApplicationId `
            -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 
    }
    catch 
    {
        if (!$servicePrincipalConnection)
        {
            $ErrorMessage = "Connection $connectionName not found."
            throw $ErrorMessage
        } 
        else
        {
            Write-Error -Message $_.Exception
            throw $_.Exception
        }
    }
    
    # Hard code name of resource group
    $resourceGroupName = '<resourceGroupName>' 
    $VMs = Get-AzureRmVM -ResourceGroupName $resourceGroupName

    Write-Output "All the VMs in the $resourceGroupName resource group in the VM deallocated state will be started"
    Write-Output ''
    Write-Output 'This includes the following VMs:'
    $VMs | % {$_.Name}
    Write-Output ''

    foreach -parallel ($VM in $VMs)
    {
        $vmName = $VM.Name
        $powerStateStatus = (Get-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -Status).Statuses[1].DisplayStatus
        if($powerStateStatus -like 'VM deallocated')
        {
            $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
            $output = $time + " INFO: Starting VM $vmName."
            Write-Output $output
            $null = Start-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name
            $powerStateStatus = (Get-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -Status).Statuses[1].DisplayStatus

            if($powerStateStatus -like 'VM running')
            {
                $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
                $output = $time + " INFO: $vmName started."
                Write-Output $output
            }

            else
            {
                $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
                $output = $time + " INFO: $vmName failed to start."
                Write-Output $output         
            }         
        }

        else
        {
            $time = Get-Date -Format MM/dd/yyy_hh:mm:ss
            $output = $time + " INFO: $vmName is already started."
            Write-Output $output         
        }
    }
}
