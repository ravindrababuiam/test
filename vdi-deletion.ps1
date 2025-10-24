<#
.SYNOPSIS
    Lists all resources associated with an AVD session host.
.DESCRIPTION
    This runbook identifies and lists all Azure resources associated with an Azure Virtual Desktop (AVD) session host.
    It requires only the session host name and will automatically use credentials stored in the Automation Account.
.PARAMETER SessionHostName
    The name of the AVD session host VM.
.PARAMETER ResourceGroupName
    (Optional) The resource group containing the AVD session host VM. If not provided, the script will attempt to find it.
.PARAMETER TenantId
    (Optional) The Azure AD tenant ID for authentication. If not provided, will try to get from Automation Account.
.PARAMETER ApplicationId
    (Optional) The Service Principal Application (Client) ID. If not provided, will try to get from Automation Account.
.PARAMETER ApplicationSecret
    (Optional) The Service Principal secret. If not provided, will try to get from Automation Account.
.PARAMETER CompactOutput
    (Optional) Use a more compact output format.
.PARAMETER SkipGraphAPI
    (Optional) Skip Microsoft Graph API calls for Entra ID and Intune information.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SessionHostName,
    
    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ApplicationId,
    
    [Parameter(Mandatory = $false)]
    [string]$ApplicationSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$CompactOutput,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipGraphAPI,

    [Parameter(Mandatory = $false)]
    [bool]$Decommission = $true,

    [Parameter(Mandatory = $false)]
    [bool]$ConfirmDecommission = $true,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAvdRemoval = $false
)

# Add System.Web assembly for URL encoding
Add-Type -AssemblyName System.Web

function Write-TableOutput {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Properties
    )
    
    Write-Output ""
    Write-Output "╔════════════════════════════════════════════════════════════════╗"
    Write-Output "║ $($Title.PadRight(60)) ║"
    Write-Output "╚════════════════════════════════════════════════════════════════╝"
    
    if (-not $Properties -or $Properties.Count -eq 0) {
        # If no properties specified, use all properties from the first item
        if ($Data -and $Data.Count -gt 0 -and $Data[0] -ne $null) {
            $Properties = $Data[0].PSObject.Properties.Name
        }
    }
    
    if ($Data -and $Data.Count -gt 0 -and $Properties -and $Properties.Count -gt 0) {
        # Determine column widths (minimum 15 characters, maximum 30)
        $columnWidths = @{}
        foreach ($prop in $Properties) {
            $maxLength = [Math]::Max(15, [Math]::Min(30, ($prop.Length)))
            foreach ($item in $Data) {
                if ($item.$prop) {
                    $valueLength = "$($item.$prop)".Length
                    $maxLength = [Math]::Max($maxLength, [Math]::Min(30, $valueLength))
                }
            }
            $columnWidths[$prop] = $maxLength
        }
        
        # Print header
        $headerLine = "| "
        $separatorLine = "|-"
        foreach ($prop in $Properties) {
            $headerLine += "$($prop.PadRight($columnWidths[$prop])) | "
            $separatorLine += "$("-" * $columnWidths[$prop])-|-"
        }
        Write-Output $headerLine
        Write-Output $separatorLine
        
        # Print data rows
        foreach ($item in $Data) {
            $line = "| "
            foreach ($prop in $Properties) {
                $value = if ($null -eq $item.$prop) { "" } else { "$($item.$prop)" }
                # Truncate if too long
                if ($value.Length -gt $columnWidths[$prop]) {
                    $value = $value.Substring(0, $columnWidths[$prop] - 3) + "..."
                }
                $line += "$($value.PadRight($columnWidths[$prop])) | "
            }
            Write-Output $line
        }
        
        # Print table footer
        $footerLine = "|"
        foreach ($prop in $Properties) {
            $footerLine += "$("-" * ($columnWidths[$prop] + 2))|"
        }
        Write-Output $footerLine
        Write-Output "Total count: $($Data.Count) item(s)"
    }
    else {
        Write-Output "No data available."
    }
}

# Add this function to display resources in specified order
function Write-OrderedInventorySummary {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$InventoryResults
    )
    
    Write-Output "`n============================================================="
    Write-Output "           RESOURCES TO BE DECOMMISSIONED                    "
    Write-Output "============================================================="
    
    # 1. AVD Host Pool Registrations
    if ($InventoryResults.AVD.SessionHosts -and $InventoryResults.AVD.SessionHosts.Count -gt 0) {
        Write-Output "`n[1. AVD HOST POOL REGISTRATIONS]"
        foreach ($sessionHost in $InventoryResults.AVD.SessionHosts) {
            Write-Output "Session Host: $($sessionHost.Name)"
            Write-Output "Host Pool: $($sessionHost.HostPoolName)"
            Write-Output "Status: $($sessionHost.Status)"
            if ($sessionHost.AssignedUser) { 
                Write-Output "Assigned User: $($sessionHost.AssignedUser)" 
            }
        }
    } else {
        Write-Output "`n[1. AVD HOST POOL REGISTRATIONS] - None found"
    }
    
    # 2. Intune Enrollment
    if ($InventoryResults.Intune.DeviceId) {
        Write-Output "`n[2. INTUNE ENROLLMENT]"
        Write-Output "Device ID: $($InventoryResults.Intune.DeviceId)"
        Write-Output "Enrolled: $($InventoryResults.Intune.Enrolled)"
        if ($InventoryResults.Intune.EnrollmentDate) {
            Write-Output "Enrollment Date: $($InventoryResults.Intune.EnrollmentDate)"
        }
    } else {
        Write-Output "`n[2. INTUNE ENROLLMENT] - None found"
    }
    
    # 3. EntraID Enrollment
    if ($InventoryResults.EntraID.DeviceId) {
        Write-Output "`n[3. ENTRA ID ENROLLMENT]"
        Write-Output "Device ID: $($InventoryResults.EntraID.DeviceId)"
        Write-Output "Join Type: $($InventoryResults.EntraID.JoinType)"
        Write-Output "MDM Enrolled: $($InventoryResults.EntraID.MDMEnrolled)"
    } else {
        Write-Output "`n[3. ENTRA ID ENROLLMENT] - None found"
    }
    
    # 4. Virtual Machine
    Write-Output "`n[4. VIRTUAL MACHINE]"
    Write-Output "Name: $($InventoryResults.SessionHost.Name)"
    Write-Output "Size: $($InventoryResults.SessionHost.Size)"
    Write-Output "Power State: $($InventoryResults.SessionHost.PowerState)"
    Write-Output "Resource Group: $($InventoryResults.SessionHost.Id.Split('/')[4])"
    
    # 5. Network Interfaces
    if ($InventoryResults.NetworkInterfaces -and $InventoryResults.NetworkInterfaces.Count -gt 0) {
        Write-Output "`n[5. NETWORK INTERFACES]"
        foreach ($nic in $InventoryResults.NetworkInterfaces) {
            Write-Output "Name: $($nic.Name)"
            Write-Output "IP Address: $($nic.PrivateIpAddress)"
        }
    } else {
        Write-Output "`n[5. NETWORK INTERFACES] - None found"
    }
    
    # 6. Managed Identity
    if ($InventoryResults.ManagedIdentity -and $InventoryResults.ManagedIdentity.Count -gt 0) {
        Write-Output "`n[6. MANAGED IDENTITY]"
        foreach ($identity in $InventoryResults.ManagedIdentity) {
            Write-Output "Name: $($identity.Name)"
            if ($identity.PrincipalId) {
                Write-Output "Principal ID: $($identity.PrincipalId)"
            }
        }
    } else {
        Write-Output "`n[6. MANAGED IDENTITY] - None found"
    }
    
    # 7. Disks
    if ($InventoryResults.Disks -and $InventoryResults.Disks.Count -gt 0) {
        Write-Output "`n[7. DISKS]"
        foreach ($disk in $InventoryResults.Disks) {
            Write-Output "Name: $($disk.Name)"
            Write-Output "Type: $($disk.DiskType)"
            Write-Output "Size: $($disk.SizeGB) GB"
            Write-Output "State: $($disk.DiskState)"
            Write-Output "------------------------"
        }
    } else {
        Write-Output "`n[7. DISKS] - None found"
    }
    
    # 8. Disk Encryption Set
    if ($InventoryResults.DiskEncryptionSet -and $InventoryResults.DiskEncryptionSet.Count -gt 0) {
        Write-Output "`n[8. DISK ENCRYPTION SET]"
        foreach ($des in $InventoryResults.DiskEncryptionSet) {
            Write-Output "Name: $($des.Name)"
            if ($des.EncryptionType) {
                Write-Output "Encryption Type: $($des.EncryptionType)"
            }
        }
    } else {
        Write-Output "`n[8. DISK ENCRYPTION SET] - None found"
    }
    
    # 9. Key Vault
    if ($InventoryResults.KeyVault -and $InventoryResults.KeyVault.Count -gt 0) {
        Write-Output "`n[9. KEY VAULT]"
        foreach ($kv in $InventoryResults.KeyVault) {
            Write-Output "Name: $($kv.Name)"
            if ($kv.VaultUri) {
                Write-Output "URI: $($kv.VaultUri)"
            }
            Write-Output "Resource Group: $($kv.ResourceGroup)"
        }
    } else {
        Write-Output "`n[9. KEY VAULT] - None found"
    }
    
    Write-Output "`n============================================================="
}

function Decommission-AVDSessionHost {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$InventoryResults,
        
        [Parameter(Mandatory = $false)]
        [switch]$WhatIf = $false
    )
    
    # Validate that inventory data is available
    if (-not $InventoryResults -or -not $InventoryResults.SessionHost) {
        Write-Error "No valid inventory results provided. Run the inventory first."
        return
    }
    
    # Retrieve session host information
    $vmName = $InventoryResults.SessionHost.Name
    $resourceGroup = $InventoryResults.SessionHost.Id.Split('/')[4]
    
    Write-Output ""
    Write-Output "============================================================="
    Write-Output "           AVD SESSION HOST DECOMMISSIONING"
    Write-Output "============================================================="
    Write-Output "Session Host: $vmName"
    Write-Output "Resource Group: $resourceGroup"
    Write-Output "Subscription: $((Get-AzContext).Subscription.Name)"
    Write-Output "Subscription ID: $((Get-AzContext).Subscription.Id)"
    Write-Output "Mode: $(if($WhatIf){'SIMULATION - NO CHANGES WILL BE MADE'}else{'ACTUAL DELETION'})"
    Write-Output "============================================================="
    Write-Output ""
    
    # Function to report status
    function Write-DecommStatus {
        param (
            [string]$Resource,
            [string]$Status,
            [string]$Message
        )
        
        switch ($Status) {
            "Success" { $statusText = "SUCCESS" }
            "Skipped" { $statusText = "SKIPPED" }
            "Failed"  { $statusText = "FAILED" }
            "Info"    { $statusText = "INFO" }
            default   { $statusText = $Status }
        }
        
        Write-Output "[$Resource] $statusText - $Message"
    }
    
    Write-Output "DECOMMISSIONING SEQUENCE STARTED: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
      
    # STEP 1: Remove from AVD Host Pool

    # STEP 1: Remove from AVD Host Pool(modified)
if (-not $SkipAvdRemoval) {
    Write-Output "`nSTEP 1: Removing from AVD Host Pool registrations..."
    
    foreach ($sessionHost in $InventoryResults.AVD.SessionHosts) {
        $hostPoolName = $sessionHost.HostPoolName
        $hostPoolRg = ($InventoryResults.AVD.HostPools | Where-Object { $_.Name -eq $hostPoolName }).ResourceGroup
        
        # Make sure we're using the exact session host name as it appears in the host pool
        # The session host name in AVD is in the format 'hostname.domain.com'
        $sessionHostName = $sessionHost.Name
        
        # Extract the actual VM name and FQDN parts if needed
        if ($sessionHostName -match '/(.+)$') {
            # The name is already in the format 'hostpoolname/hostname.domain.com'
            $sessionHostName = $matches[1]
        }
        
        Write-Output "Working with Host Pool: $hostPoolName, Resource Group: $hostPoolRg"
        Write-Output "Session Host to remove: $sessionHostName"
        
        try {
            if (-not $WhatIf) {
                # First check for active sessions - if found, halt execution
                Write-Output "Checking for active sessions..."
                try {
                    $sessions = Get-AzWvdUserSession -HostPoolName $hostPoolName -ResourceGroupName $hostPoolRg -SessionHostName $sessionHostName -ErrorAction SilentlyContinue
                    if ($sessions -and $sessions.Count -gt 0) {
                        Write-Output "HALT: Found $($sessions.Count) active sessions."
                        Write-Output "Active sessions:"
                        foreach ($session in $sessions) {
                            Write-Output "  User: $($session.UserPrincipalName), Session ID: $($session.Id)"
                            Write-Output "  Connected: $($session.SessionState)"
                            Write-Output "  ------------"
                        }
                        Write-Output "Decommissioning halted due to active sessions. Please terminate sessions before proceeding."
                        
                        # Terminate the runbook/script execution
                        throw "Decommissioning halted due to active sessions. Please terminate sessions before proceeding."
                    } else {
                        Write-Output "No active sessions found. Proceeding with decommissioning."
                    }
                } catch {
                    if ($_.Exception.Message -like "*active sessions*") {
                        # This is our custom exception - rethrow it to terminate script
                        throw $_
                    }
                    Write-Warning "Error checking for user sessions: $_"
                }
                
                # Drain mode to prevent new connections
                Write-Output "Setting session host to drain mode: $sessionHostName"
                Update-AzWvdSessionHost -ResourceGroupName $hostPoolRg -HostPoolName $hostPoolName -Name $sessionHostName -AllowNewSession:$false -ErrorAction Continue
                
                # Wait a moment for the drain mode to apply
                Start-Sleep -Seconds 2
                
                # Remove session host from host pool with Force
                Write-Output "Removing session host from host pool: $sessionHostName"
                Remove-AzWvdSessionHost -ResourceGroupName $hostPoolRg -HostPoolName $hostPoolName -Name $sessionHostName -Force -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "AVD Host Pool" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Removed $sessionHostName from host pool $hostPoolName"
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            # Check if this is our custom exception for active sessions
            if ($errorMsg -like "*active sessions*") {
                # This is our custom exception - rethrow it to terminate script
                throw $_
            }
            
            Write-DecommStatus -Resource "AVD Host Pool" -Status "Failed" -Message "Could not remove from host pool - $errorMsg"
            
            # Try an alternative approach if the first method fails
            try {
                if (-not $WhatIf) {
                    # Sometimes we need to use the full resource path format
                    Write-Output "Trying alternative approach with full resource name..."
                    $fullSessionHostName = "$hostPoolName/$sessionHostName"
                    
                    # Try with full name
                    Write-Output "Removing with full resource name: $fullSessionHostName"
                    Remove-AzWvdSessionHost -ResourceGroupName $hostPoolRg -HostPoolName $hostPoolName -Name $fullSessionHostName -Force -ErrorAction Stop
                    
                    Write-DecommStatus -Resource "AVD Host Pool" -Status "Success" -Message "Removed $sessionHostName using alternative method"
                }
            }
            catch {
                $altErrorMsg = $_.Exception.Message
                Write-DecommStatus -Resource "AVD Host Pool" -Status "Failed" -Message "Both removal methods failed - $altErrorMsg"
            }
        }
    }
}
else {
    Write-Output "`nSTEP 1: Skipping AVD Host Pool removal as requested"
}

    # STEP 2: Remove from Intune
    if ($InventoryResults.Intune.DeviceId) {
        Write-Output "`nSTEP 2: Removing from Intune enrollment..."
        $intuneDeviceId = $InventoryResults.Intune.DeviceId
        
        try {
            if (-not $WhatIf) {
                # Using Microsoft Graph REST API with the existing token
                $graphToken = $global:graphToken
                
                if ($graphToken) {
                    $graphHeaders = @{
                        "Authorization" = "Bearer $graphToken"
                        "Content-Type" = "application/json"
                    }
                    
                    $intuneUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$intuneDeviceId"
                    
                    # Delete the Intune device
                    Invoke-RestMethod -Uri $intuneUrl -Headers $graphHeaders -Method DELETE -ErrorAction Stop
                    Write-DecommStatus -Resource "Intune" -Status "Success" -Message "Removed device ID: $intuneDeviceId"
                }
                else {
                    Write-DecommStatus -Resource "Intune" -Status "Skipped" -Message "No Graph API token available, manual removal required for device ID: $intuneDeviceId"
                }
            }
            else {
                Write-DecommStatus -Resource "Intune" -Status "Skipped" -Message "Would remove device ID: $intuneDeviceId"
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Intune" -Status "Failed" -Message "Could not remove from Intune - $errorMsg"
        }
    }
    else {
        Write-Output "`nSTEP 2: No Intune enrollment found"
    }
    
    # STEP 3: Remove from EntraID
    if ($InventoryResults.EntraID.DeviceId) {
        Write-Output "`nSTEP 3: Removing from EntraID enrollment..."
        $entraDeviceId = $InventoryResults.EntraID.DeviceId
        
        try {
            if (-not $WhatIf) {
                # Using Microsoft Graph REST API with the existing token
                $graphToken = $global:graphToken
                
                if ($graphToken) {
                    $graphHeaders = @{
                        "Authorization" = "Bearer $graphToken"
                        "Content-Type" = "application/json"
                    }
                    
                    $entraUrl = "https://graph.microsoft.com/v1.0/devices/$entraDeviceId"
                    
                    # Delete the EntraID device
                    Invoke-RestMethod -Uri $entraUrl -Headers $graphHeaders -Method DELETE -ErrorAction Stop
                    Write-DecommStatus -Resource "EntraID" -Status "Success" -Message "Removed device ID: $entraDeviceId"
                }
                else {
                    Write-DecommStatus -Resource "EntraID" -Status "Skipped" -Message "No Graph API token available, manual removal required for device ID: $entraDeviceId"
                }
            }
            else {
                Write-DecommStatus -Resource "EntraID" -Status "Skipped" -Message "Would remove device ID: $entraDeviceId"
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "EntraID" -Status "Failed" -Message "Could not remove from EntraID - $errorMsg"
        }
    }
    else {
        Write-Output "`nSTEP 3: No EntraID enrollment found"
    }
    
    # STEP 4: Delete the Virtual Machine
    Write-Output "`nSTEP 4: Deleting Virtual Machine..."
    try {
        if (-not $WhatIf) {
            # Stop the VM first if it's running
            $vmStatus = Get-AzVM -ResourceGroupName $resourceGroup -Name $vmName -Status -ErrorAction SilentlyContinue
            $powerState = ($vmStatus.Statuses | Where-Object { $_.Code -match 'PowerState/' }).Code -replace 'PowerState/', ''
            
            if ($powerState -ne "deallocated") {
                Write-Output "Stopping VM $vmName..."
                Stop-AzVM -ResourceGroupName $resourceGroup -Name $vmName -Force -ErrorAction Stop
            }
            else {
                Write-Output "VM $vmName is already stopped"
            }
            
            # Delete the VM
            Write-Output "Deleting VM $vmName..."
            Remove-AzVM -ResourceGroupName $resourceGroup -Name $vmName -Force -ErrorAction Stop
        }
        
        Write-DecommStatus -Resource "Virtual Machine" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted VM: $vmName"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DecommStatus -Resource "Virtual Machine" -Status "Failed" -Message "Could not delete VM - $errorMsg"
    }
    
    # STEP 5: Delete Network Interfaces
    Write-Output "`nSTEP 5: Deleting Network Interfaces..."
    foreach ($nic in $InventoryResults.NetworkInterfaces) {
        try {
            if (-not $WhatIf) {
                Write-Output "Deleting NIC: $($nic.Name)..."
                Remove-AzNetworkInterface -ResourceGroupName $resourceGroup -Name $nic.Name -Force -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "Network Interface" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted NIC: $($nic.Name)"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Network Interface" -Status "Failed" -Message "Could not delete NIC $($nic.Name) - $errorMsg"
        }
    }
    
    # STEP 6: Delete Managed Identity
    Write-Output "`nSTEP 6: Deleting Managed Identities..."
    foreach ($identity in $InventoryResults.ManagedIdentity) {
        try {
            if (-not $WhatIf) {
                Write-Output "Deleting Managed Identity: $($identity.Name)..."
                Remove-AzUserAssignedIdentity -ResourceGroupName $identity.ResourceGroup -Name $identity.Name  -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "Managed Identity" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted identity: $($identity.Name)"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Managed Identity" -Status "Failed" -Message "Could not delete identity $($identity.Name) - $errorMsg"
        }
    }
    
    # STEP 7: Delete all Disks
    Write-Output "`nSTEP 7: Deleting Disks..."
    foreach ($disk in $InventoryResults.Disks) {
        $diskName = $disk.Name
        $diskType = $disk.DiskType
        
        try {
            if (-not $WhatIf) {
                Write-Output "Deleting $diskType disk: $diskName..."
                # Wait a bit to ensure the VM deletion is complete and disk is ready to delete
                Start-Sleep -Seconds 5
                Remove-AzDisk -ResourceGroupName $resourceGroup -DiskName $diskName -Force -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "Disk" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted $diskType disk: $diskName"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Disk" -Status "Failed" -Message "Could not delete disk $diskName - $errorMsg"
        }
    }
    
    # STEP 8: Delete Disk Encryption Set
    Write-Output "`nSTEP 8: Deleting Disk Encryption Sets..."
    foreach ($des in $InventoryResults.DiskEncryptionSet) {
        try {
            if (-not $WhatIf) {
                Write-Output "Deleting Disk Encryption Set: $($des.Name)..."
                Remove-AzDiskEncryptionSet -ResourceGroupName $des.ResourceGroup -Name $des.Name -Force -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "Disk Encryption Set" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted DES: $($des.Name)"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Disk Encryption Set" -Status "Failed" -Message "Could not delete DES $($des.Name) - $errorMsg"
        }
    }
    
    # STEP 9: Delete Key Vault
    Write-Output "`nSTEP 9: Deleting Key Vaults..."
    foreach ($kv in $InventoryResults.KeyVault) {
        try {
            if (-not $WhatIf) {
                Write-Output "Deleting Key Vault: $($kv.Name) - WARNING: This might contain other secrets!"
                Remove-AzKeyVault -ResourceGroupName $kv.ResourceGroup -Name $kv.Name -Force -ErrorAction Stop
            }
            
            Write-DecommStatus -Resource "Key Vault" -Status $(if($WhatIf){"Skipped"}else{"Success"}) -Message "Deleted Key Vault: $($kv.Name)"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-DecommStatus -Resource "Key Vault" -Status "Failed" -Message "Could not delete Key Vault $($kv.Name) - $errorMsg"
        }
    }
    
    Write-Output ""
    Write-Output "============================================================="
    Write-Output "           DECOMMISSIONING PROCESS COMPLETE"
    Write-Output "============================================================="
    Write-Output "Completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    if ($WhatIf) {
        Write-Output "This was a simulation. No resources were actually deleted."
        Write-Output "Run again with -Decommission -ConfirmDecommission to perform actual deletion."
    }
    else {
        Write-Output "All specified resources have been successfully processed for decommision."
    }
    Write-Output "============================================================="
    Write-Output ""
}

try {
    # Check if we need to authenticate with Service Principal
    $needToAuthenticate = $true
    
    # Check if we're already authenticated
    try {
        $azContext = Get-AzContext -ErrorAction Stop
        if ($azContext) {
            Write-Output "Already authenticated as $($azContext.Account.Id) in tenant $($azContext.Tenant.Id)"
            $needToAuthenticate = $false
            
            # If we're in an Automation Account, we still need to use the provided Service Principal
            if ($PSPrivateMetadata.JobId) {
                $needToAuthenticate = $true
            }
        }
    }
    catch {
        $needToAuthenticate = $true
    }
    
    # Authenticate if needed
    if ($needToAuthenticate) {
        # Check if we're running in an Automation Account
        if ($PSPrivateMetadata.JobId) {
            # We're in an Automation Account - try to get credentials from variables if not provided
            Write-Output "Running in Azure Automation"
            
            if (-not $TenantId) {
                try {
                    # Try standard variable first
                    $TenantId = Get-AutomationVariable -Name "TenantId" -ErrorAction Stop
                    Write-Output "Retrieved TenantId from Automation variable 'TenantId'"
                }
                catch {
                    try {
                        # Try your specific variable name
                        $TenantId = Get-AutomationVariable -Name "AZURE_TENANT_ID" -ErrorAction Stop
                        Write-Output "Retrieved TenantId from Automation variable 'AZURE_TENANT_ID'"
                    }
                    catch {
                        Write-Warning "TenantId not found in Automation variables. Will try to continue with current context."
                    }
                }
            }
            
            if (-not $ApplicationId -or -not $ApplicationSecret) {
                try {
                    # First try to get from credential object
                    $servicePrincipalCred = Get-AutomationPSCredential -Name "AVDResourcesSP" -ErrorAction Stop
                    $ApplicationId = $servicePrincipalCred.UserName
                    $ApplicationSecret = $servicePrincipalCred.GetNetworkCredential().Password
                    Write-Output "Retrieved Service Principal credentials from Automation credential 'AVDResourcesSP'"
                }
                catch {
                    # If credential not found, try individual variables
                    try {
                        # Try standard variable names first
                        $ApplicationId = Get-AutomationVariable -Name "ApplicationId" -ErrorAction Stop
                        Write-Output "Retrieved ApplicationId from Automation variable 'ApplicationId'"
                        
                        $ApplicationSecret = Get-AutomationVariable -Name "ApplicationSecret" -ErrorAction Stop
                        Write-Output "Retrieved ApplicationSecret from Automation variable 'ApplicationSecret'"
                    }
                    catch {
                        try {
                            # Try your specific variable names
                            $ApplicationId = Get-AutomationVariable -Name "AZURE_CLIENT_ID" -ErrorAction Stop
                            Write-Output "Retrieved ApplicationId from Automation variable 'AZURE_CLIENT_ID'"
                            
                            $ApplicationSecret = Get-AutomationVariable -Name "AZURE_CLIENT_SECRET" -ErrorAction Stop
                            Write-Output "Retrieved ApplicationSecret from Automation variable 'AZURE_CLIENT_SECRET'"
                        }
                        catch {
                            Write-Warning "Service Principal credentials not found in Automation variables. Will try to continue with current context."
                        }
                    }
                }
            }
            
            # Get subscription ID if available
            try {
                $subscriptionId = Get-AutomationVariable -Name "AZURE_SUBSCRIPTION_ID" -ErrorAction Stop
                Write-Output "Retrieved SubscriptionId from Automation variable 'AZURE_SUBSCRIPTION_ID'"
            }
            catch {
                Write-Warning "SubscriptionId not found in Automation variables. Will use current subscription context."
            }
        }
        
        # Check if we have the required authentication parameters
        if (-not $TenantId -or -not $ApplicationId -or -not $ApplicationSecret) {
            throw "Authentication parameters missing. Please specify TenantId, ApplicationId, and ApplicationSecret."
        }
        
        # Log in with Service Principal
        $SecurePassword = ConvertTo-SecureString -String $ApplicationSecret -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $SecurePassword
        
        # Connect using the Service Principal
        Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential
        
        # Set subscription context if we have a subscription ID
        if ($subscriptionId) {
            Set-AzContext -SubscriptionId $subscriptionId
            Write-Output "Set context to subscription ID: $subscriptionId"
        }
        
        Write-Output "Successfully authenticated with Service Principal"
    }
# Extract subscription identifier from session hostname (4th and 5th characters)
$subscriptionIdentifier = $null
if ($SessionHostName.Length -ge 5) {
    $subscriptionIdentifier = $SessionHostName.Substring(3, 2)
    Write-Output "Extracted subscription identifier from hostname: $subscriptionIdentifier"
}
# Map the identifier to corresponding subscription variable name
$subscriptionVarName = $null
switch ($subscriptionIdentifier) {
    "S1" { $subscriptionVarName = "Subscription1" }
    "S2" { $subscriptionVarName = "Subscription2" }
    "S3" { $subscriptionVarName = "Subscription3" }
    "S4" { $subscriptionVarName = "Subscription4" }
    default {
        Write-Warning "Session host name does not contain a recognized subscription identifier (S1, S2, S3, or S4). Using current subscription context."
    }
}

# Get subscription ID from Automation Account variable if identifier was found
$targetSubscriptionId = $null
if ($subscriptionVarName -and $PSPrivateMetadata.JobId) {  # Check if we're in Automation Account
    try {
        $targetSubscriptionId = Get-AutomationVariable -Name $subscriptionVarName -ErrorAction Stop
        Write-Output "Retrieved subscription ID from Automation variable '$subscriptionVarName'"
    }
    catch {
        Write-Warning "Could not retrieve subscription ID from variable '$subscriptionVarName': $_"
    }
}

# Set Azure context to target subscription if found
if ($targetSubscriptionId) {
    try {
        Write-Output "Setting context to subscription ID: $targetSubscriptionId"
        Set-AzContext -SubscriptionId $targetSubscriptionId -ErrorAction Stop
        Write-Output "Successfully set subscription context"
    }
    catch {
        Write-Warning "Failed to set subscription context: $_"
        # Continue with current context as fallback
    }
}
# Find the VM across all resource groups if ResourceGroupName not provided
if (-not $ResourceGroupName) {
    Write-Output "Resource group not specified. Searching for VM $SessionHostName in subscription $((Get-AzContext).Subscription.Name)..."
    $vm = Get-AzVM -Name $SessionHostName -ErrorAction SilentlyContinue
    if (-not $vm) {
        throw "VM $SessionHostName not found in subscription $((Get-AzContext).Subscription.Name). Please check the VM name or specify a resource group."
    }
    $ResourceGroupName = $vm.ResourceGroupName
    Write-Output "Found VM in resource group: $ResourceGroupName"
}
else {
    # Get the VM details
    $vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $SessionHostName -ErrorAction Stop
    Write-Output "Found VM: $($vm.Name) in resource group: $ResourceGroupName, subscription: $((Get-AzContext).Subscription.Name)"
}
    
    # Initialize results object
    $results = [PSCustomObject]@{
        SessionHost = [PSCustomObject]@{
            Name = $vm.Name
            Id = $vm.Id
            Location = $vm.Location
            Size = $vm.HardwareProfile.VmSize
            OSType = $vm.StorageProfile.OsDisk.OsType
            ProvisioningState = $vm.ProvisioningState
            PowerState = $null
            LastPowerOnTime = $null
            LastLoggedInTime = $null
            AssignedUser = $null
        }
        NetworkInterfaces = @()
        Disks = @()
        Extensions = @()
        ManagedIdentity = @()
        KeyVault = @()
        DiskEncryptionSet = @()
        AVD = [PSCustomObject]@{
            HostPools = @()
            SessionHosts = @()
        }
        EntraID = [PSCustomObject]@{
            DeviceId = $null
            JoinType = $null
            MDMEnrolled = $null
            LastSyncTime = $null
            TrustType = $null
            EnrollmentDate = $null
        }
        Intune = [PSCustomObject]@{
            DeviceId = $null
            Enrolled = $false
            EnrollmentDate = $null
            LastContactTime = $null
            ComplianceState = $null
            ManagementAgent = $null
        }
        RelatedResources = @()
    }
    
    # Get VM power state and activity information
    try {
        Write-Output "Getting VM power state and activity information..."
        $vmStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $SessionHostName -Status
        if ($vmStatus) {
            # Get current power state
            $powerState = ($vmStatus.Statuses | Where-Object { $_.Code -match 'PowerState/' }).Code -replace 'PowerState/', ''
            $results.SessionHost.PowerState = $powerState
            Write-Output "VM Power State: $powerState"
            
            # Try to get activity log for boot events
            try {
                $startTime = (Get-Date).AddDays(-30) # Look back 30 days
                $endTime = Get-Date
                
                # Look for start/deallocate events in activity log
                $activityLogs = Get-AzActivityLog -ResourceId $vm.Id -StartTime $startTime -EndTime $endTime -WarningAction SilentlyContinue
                
                # Find the most recent start event (VM powered on)
                $startEvent = $activityLogs | Where-Object { $_.OperationName.Value -eq 'Microsoft.Compute/virtualMachines/start/action' -and $_.Status.Value -eq 'Succeeded' } | Sort-Object -Property EventTimestamp -Descending | Select-Object -First 1
                if ($startEvent) {
                    $results.SessionHost.LastPowerOnTime = $startEvent.EventTimestamp
                    Write-Output "Last Power On Time: $($startEvent.EventTimestamp)"
                }
            }
            catch {
                Write-Warning "Could not retrieve VM activity logs: $_"
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve VM power state: $_"
    }

    # Get Entra ID device information using Microsoft Graph REST API
    if (-not $SkipGraphAPI) {
        try {
            Write-Output "Checking for Entra ID device enrollment using Graph API..."
            
            # Use the same Service Principal to get a token for Microsoft Graph
            $graphTokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            $graphScope = "https://graph.microsoft.com/.default"
            
            $graphTokenBody = @{
                client_id     = $ApplicationId
                client_secret = $ApplicationSecret
                scope         = $graphScope
                grant_type    = "client_credentials"
            }
            
            # Get Graph API token
            Write-Output "Requesting Microsoft Graph API token..."
            $graphTokenResponse = Invoke-RestMethod -Uri $graphTokenUrl -Method POST -Body $graphTokenBody -ErrorAction Stop
            $graphToken = $graphTokenResponse.access_token
            
            if ($graphToken) {
                Write-Output "Successfully acquired token for Microsoft Graph API"
                
                # Set up headers for Graph API calls
                $graphHeaders = @{
                    "Authorization" = "Bearer $graphToken"
                    "Content-Type"  = "application/json"
                }
                
                # Query for the device by display name
                $deviceFilter = [System.Web.HttpUtility]::UrlEncode("displayName eq '$SessionHostName'")
                $deviceUrl = "https://graph.microsoft.com/v1.0/devices?`$filter=$deviceFilter"
                
                Write-Output "Querying Microsoft Graph API for device: $SessionHostName"
                $deviceResponse = Invoke-RestMethod -Uri $deviceUrl -Headers $graphHeaders -Method GET -ErrorAction Stop
                
                if ($deviceResponse.value -and $deviceResponse.value.Count -gt 0) {
                    $device = $deviceResponse.value[0]
                    $results.EntraID.DeviceId = $device.id
                    $results.EntraID.JoinType = $device.joinType
                    $results.EntraID.MDMEnrolled = $device.mdmAppId -ne $null
                    $results.EntraID.LastSyncTime = $device.approximateLastSignInDateTime
                    $results.EntraID.TrustType = $device.trustType
                    $results.EntraID.EnrollmentDate = $device.registrationDateTime
                    
                    Write-Output "Found Entra ID device: $($device.id)"
                    Write-Output "  Join Type: $($device.joinType)"
                    Write-Output "  MDM Enrolled: $($results.EntraID.MDMEnrolled)"
                    Write-Output "  Last Sync Time: $($device.approximateLastSignInDateTime)"
                }
                else {
                    Write-Output "No Entra ID device record found with name: $SessionHostName"
                }
                
                # Now query for Intune device information
                try {
                    Write-Output "Checking for Intune device enrollment using Graph API..."
                    
                    $intuneFilter = [System.Web.HttpUtility]::UrlEncode("deviceName eq '$SessionHostName'")
                    $intuneUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$intuneFilter"
                    
                    Write-Output "Querying Microsoft Graph API for Intune device: $SessionHostName"
                    $intuneResponse = Invoke-RestMethod -Uri $intuneUrl -Headers $graphHeaders -Method GET -ErrorAction Stop
                    
                    if ($intuneResponse.value -and $intuneResponse.value.Count -gt 0) {
                        $intuneDevice = $intuneResponse.value[0]
                        $results.Intune.DeviceId = $intuneDevice.id
                        $results.Intune.Enrolled = $true
                        $results.Intune.EnrollmentDate = $intuneDevice.enrolledDateTime
                        $results.Intune.LastContactTime = $intuneDevice.lastSyncDateTime
                        $results.Intune.ComplianceState = $intuneDevice.complianceState
                        $results.Intune.ManagementAgent = $intuneDevice.managementAgent
                        
                        Write-Output "Found Intune device: $($intuneDevice.id)"
                        Write-Output "  Enrollment Date: $($intuneDevice.enrolledDateTime)"
                        Write-Output "  Last Contact Time: $($intuneDevice.lastSyncDateTime)"
                        Write-Output "  Compliance State: $($intuneDevice.complianceState)"
                    }
                    else {
                        Write-Output "No Intune device record found with name: $SessionHostName"
                    }
                }
                catch {
                    Write-Warning "Error retrieving Intune device information: $_"
                    Write-Warning "Response: $($_.Exception.Response)"
                }
            }
            else {
                Write-Warning "Could not acquire token for Microsoft Graph API"
            }
        }
        catch {
            Write-Warning "Error checking Entra ID/Intune enrollment: $_"
            if ($_.Exception.Response) {
                Write-Warning "Status code: $($_.Exception.Response.StatusCode.value__)"
                Write-Warning "Status description: $($_.Exception.Response.StatusDescription)"
            }
        }
    }
    else {
        Write-Output "Skipping Microsoft Graph API calls for Entra ID and Intune information"
    }
    
    # Get resources with specific naming pattern based on VM name
    $vmNameBase = $SessionHostName
    Write-Output "Looking for resources with naming pattern related to $vmNameBase"
    
    # Look for Managed Identity (-p-id suffix)
    $managedIdentityName = "$vmNameBase-p-id"
    $identityResource = Get-AzResource -Name $managedIdentityName -ErrorAction SilentlyContinue
    if ($identityResource) {
        Write-Output "Found Managed Identity: $managedIdentityName"
        $results.ManagedIdentity += [PSCustomObject]@{
            Name = $identityResource.Name
            Id = $identityResource.ResourceId
            Type = $identityResource.ResourceType
            ResourceGroup = $identityResource.ResourceGroupName
        }
        
        # Try to get more details about the identity
        try {
            $userAssignedIdentity = Get-AzUserAssignedIdentity -ResourceGroupName $identityResource.ResourceGroupName -Name $identityResource.Name
            if ($userAssignedIdentity) {
                $results.ManagedIdentity[-1] | Add-Member -NotePropertyName PrincipalId -NotePropertyValue $userAssignedIdentity.PrincipalId
                $results.ManagedIdentity[-1] | Add-Member -NotePropertyName ClientId -NotePropertyValue $userAssignedIdentity.ClientId
            }
        }
        catch {
            Write-Warning "Could not get detailed identity information: $_"
        }
    }
    
    # Look for Key Vault (-p-kv suffix)
    $keyVaultName = "$vmNameBase-p-kv"
    $keyVaultResource = Get-AzResource -Name $keyVaultName -ErrorAction SilentlyContinue
    if ($keyVaultResource) {
        Write-Output "Found Key Vault: $keyVaultName"
        $results.KeyVault += [PSCustomObject]@{
            Name = $keyVaultResource.Name
            Id = $keyVaultResource.ResourceId
            Type = $keyVaultResource.ResourceType
            ResourceGroup = $keyVaultResource.ResourceGroupName
        }
        
        # Try to get more details about the key vault
        try {
            $keyVault = Get-AzKeyVault -ResourceGroupName $keyVaultResource.ResourceGroupName -VaultName $keyVaultResource.Name
            if ($keyVault) {
                $results.KeyVault[-1] | Add-Member -NotePropertyName VaultUri -NotePropertyValue $keyVault.VaultUri
                $results.KeyVault[-1] | Add-Member -NotePropertyName EnabledForDiskEncryption -NotePropertyValue $keyVault.EnabledForDiskEncryption
            }
        }
        catch {
            Write-Warning "Could not get detailed key vault information: $_"
        }
    }
    
    # Look for Disk Encryption Set (-p-des suffix)
    $desName = "$vmNameBase-p-des"
    $desResource = Get-AzResource -Name $desName -ErrorAction SilentlyContinue
    if ($desResource) {
        Write-Output "Found Disk Encryption Set: $desName"
        $results.DiskEncryptionSet += [PSCustomObject]@{
            Name = $desResource.Name
            Id = $desResource.ResourceId
            Type = $desResource.ResourceType
            ResourceGroup = $desResource.ResourceGroupName
        }
        
        # Try to get more details about the disk encryption set
        try {
            $des = Get-AzDiskEncryptionSet -ResourceGroupName $desResource.ResourceGroupName -Name $desResource.Name
            if ($des) {
                $results.DiskEncryptionSet[-1] | Add-Member -NotePropertyName EncryptionType -NotePropertyValue $des.EncryptionType
                $results.DiskEncryptionSet[-1] | Add-Member -NotePropertyName KeyVaultUrl -NotePropertyValue $des.ActiveKey.KeyUrl
            }
        }
        catch {
            Write-Warning "Could not get detailed disk encryption set information: $_"
        }
    }
    
    # Look for NIC (-nic suffix)
    $nicName = "$vmNameBase-nic"
    $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    if ($nic) {
        Write-Output "Found Network Interface: $nicName"
        $nicInfo = [PSCustomObject]@{
            Name = $nic.Name
            Id = $nic.Id
            PrivateIpAddress = $nic.IpConfigurations[0].PrivateIpAddress
            SubnetId = $nic.IpConfigurations[0].Subnet.Id
        }
        
        if ($nic.NetworkSecurityGroup) {
            $nicInfo | Add-Member -NotePropertyName NetworkSecurityGroupId -NotePropertyValue $nic.NetworkSecurityGroup.Id
        }
        
        $results.NetworkInterfaces += $nicInfo
    }

    # Look for OS Disk (matching pattern with OsDisk in name)

# Get disks directly from the VM object first to ensure we catch all attached disks
Write-Output "Getting disk information from VM..."

# Add the OS disk
Write-Output "Found OS Disk: $($vm.StorageProfile.OsDisk.Name)"
$results.Disks += [PSCustomObject]@{
    Name = $vm.StorageProfile.OsDisk.Name
    Id = $vm.StorageProfile.OsDisk.ManagedDisk.Id
    DiskType = "OS"
    SizeGB = $null  # Will be filled from Get-AzDisk
    Sku = $null     # Will be filled from Get-AzDisk
    Encryption = $null
    DiskState = "Attached"
}

# Add all data disks from the VM object
foreach ($dataDisk in $vm.StorageProfile.DataDisks) {
    Write-Output "Found Data Disk: $($dataDisk.Name)"
    $results.Disks += [PSCustomObject]@{
        Name = $dataDisk.Name
        Id = $dataDisk.ManagedDisk.Id
        DiskType = "Data"
        SizeGB = $dataDisk.DiskSizeGB
        Lun = $dataDisk.Lun
        Sku = $null  # Will be filled from Get-AzDisk
        Encryption = $null
        DiskState = "Attached"
    }
}

# Now get detailed information for all disks
$allDiskNames = @($vm.StorageProfile.OsDisk.Name) + @($vm.StorageProfile.DataDisks | ForEach-Object { $_.Name })
foreach ($diskName in $allDiskNames) {
    try {
        $diskDetails = Get-AzDisk -ResourceGroupName $ResourceGroupName -DiskName $diskName -ErrorAction SilentlyContinue
        if ($diskDetails) {
            # Find the corresponding disk in our results
            $diskInResults = $results.Disks | Where-Object { $_.Name -eq $diskName }
            if ($diskInResults) {
                # Update with detailed information
                $diskInResults.SizeGB = $diskDetails.DiskSizeGB
                $diskInResults.Sku = $diskDetails.Sku.Name
                if ($diskDetails.Encryption) {
                    $diskInResults.Encryption = $diskDetails.Encryption.Type
                }
                $diskInResults.DiskState = $diskDetails.DiskState
            }
        }
    }
    catch {
        Write-Warning "Could not get detailed information for disk '$diskName': $_"
    }
}

# Also look for any detached disks that match the VM naming pattern
# Look for OS Disk (matching pattern with OsDisk in name)
$osDiskPattern = "$vmNameBase" + "_OsDisk_"
Get-AzDisk -ResourceGroupName $ResourceGroupName | Where-Object { 
    $_.Name -like "$osDiskPattern*" -and 
    $_.Name -ne $vm.StorageProfile.OsDisk.Name 
} | ForEach-Object {
    Write-Output "Found additional OS Disk: $($_.Name)"
    $results.Disks += [PSCustomObject]@{
        Name = $_.Name
        Id = $_.Id
        DiskType = "OS"
        SizeGB = $_.DiskSizeGB
        Sku = $_.Sku.Name
        Encryption = $_.Encryption.Type
        DiskState = $_.DiskState
    }
}

# Look for Data Disks (matching pattern with DataDisk in name)
$dataDiskPattern = "$vmNameBase" + "_DataDisk_"
Get-AzDisk -ResourceGroupName $ResourceGroupName | Where-Object { 
    $_.Name -like "$dataDiskPattern*" -and
    $_.Name -notin ($vm.StorageProfile.DataDisks | ForEach-Object { $_.Name })
} | ForEach-Object {
    Write-Output "Found additional Data Disk: $($_.Name)"
    $results.Disks += [PSCustomObject]@{
        Name = $_.Name
        Id = $_.Id
        DiskType = "Data"
        SizeGB = $_.DiskSizeGB
        Sku = $_.Sku.Name
        Encryption = $_.Encryption.Type
        DiskState = $_.DiskState
    }
}    
    
    # Get VM extensions
    $extensions = Get-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $SessionHostName -ErrorAction SilentlyContinue
    foreach ($extension in $extensions) {
        Write-Output "Found VM Extension: $($extension.Name)"
        $results.Extensions += [PSCustomObject]@{
            Name = $extension.Name
            Id = $extension.Id
            Publisher = $extension.Publisher
            ExtensionType = $extension.ExtensionType
            ProvisioningState = $extension.ProvisioningState
        }
    }
    
    # Get AVD host pools and session host associations
    # Import Az.DesktopVirtualization module if available
    $moduleAvailable = Get-Module -ListAvailable -Name Az.DesktopVirtualization
    if ($moduleAvailable) {
        try {
            Import-Module Az.DesktopVirtualization -ErrorAction Stop
            
            # Get all host pools
            $hostPools = Get-AzWvdHostPool -ErrorAction Stop
            
            foreach ($hostPool in $hostPools) {
                # Get session hosts in this host pool
                $sessionHosts = Get-AzWvdSessionHost -HostPoolName $hostPool.Name -ResourceGroupName $hostPool.Id.Split('/')[4] -ErrorAction SilentlyContinue
                
                foreach ($sessionHost in $sessionHosts) {
                    # Extract VM name from session host name
                    $sessionHostVmName = $sessionHost.Name.Split('/', 2)[1].Split('.')[0]
                    
                    if ($sessionHostVmName -eq $SessionHostName) {
                        # This host pool contains our session host
                        Write-Output "Found AVD Host Pool: $($hostPool.Name)"
                        $results.AVD.HostPools += [PSCustomObject]@{
                            Name = $hostPool.Name
                            Id = $hostPool.Id
                            ResourceGroup = $hostPool.Id.Split('/')[4]
                            Type = $hostPool.HostPoolType
                            LoadBalancerType = $hostPool.LoadBalancerType
                        }
                        
                        Write-Output "Found AVD Session Host Registration: $($sessionHost.Name)"
                        $results.AVD.SessionHosts += [PSCustomObject]@{
                            Name = $sessionHost.Name
                            Id = $sessionHost.Id
                            Status = $sessionHost.Status
                            AssignedUser = $sessionHost.AssignedUser
                            HostPoolName = $hostPool.Name
                        }
                        
                        # Update the session host assigned user in the main results
                        if ($sessionHost.AssignedUser) {
                            $results.SessionHost.AssignedUser = $sessionHost.AssignedUser
                            Write-Output "Assigned User: $($sessionHost.AssignedUser)"
                        }
                        
                        # Try to get session information to determine last login time
                        try {
                            $sessions = Get-AzWvdUserSession -HostPoolName $hostPool.Name -ResourceGroupName $hostPool.Id.Split('/')[4] -SessionHostName $sessionHost.Name.Split('/', 2)[1] -ErrorAction SilentlyContinue
                            if ($sessions -and $sessions.Count -gt 0) {
                                # Get the most recent session
                                $latestSession = $sessions | Sort-Object -Property LastDisconnectTime -Descending | Select-Object -First 1
                                if ($latestSession) {
                                    $results.SessionHost.LastLoggedInTime = $latestSession.LastDisconnectTime
                                    Write-Output "Last Logged In Time: $($latestSession.LastDisconnectTime)"
                                }
                            }
                        }
                        catch {
                            Write-Warning "Could not retrieve session information: $_"
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Error accessing AVD resources: $_"
        }
    }
    else {
        Write-Warning "Az.DesktopVirtualization module not available. AVD host pool information will be limited."
        
        # Use resource graph to find related AVD resources
        if (Get-Module -ListAvailable -Name Az.ResourceGraph) {
            try {
                Import-Module Az.ResourceGraph
                
                $query = "Resources | where type =~ 'Microsoft.DesktopVirtualization/hostpools' | project id, name, resourceGroup, properties"
                $hostPools = Search-AzGraph -Query $query
                
                foreach ($hostPool in $hostPools) {
                    # Check if the host pool's VM template contains our VM name
                    if ($hostPool.properties.vmTemplate -like "*$SessionHostName*") {
                        $results.AVD.HostPools += [PSCustomObject]@{
                            Name = $hostPool.name
                            Id = $hostPool.id
                            ResourceGroup = $hostPool.resourceGroup
                        }
                    }
                }
            }
            catch {
                Write-Warning "Error querying Resource Graph: $_"
            }
        }
    }
    
    # Find other related resources using Resource Graph
    if (Get-Module -ListAvailable -Name Az.ResourceGraph) {
        try {
            Import-Module Az.ResourceGraph
            
            # Find resources that reference the VM
            $query = "Resources | where properties contains '$($vm.Id)' or id contains '$SessionHostName' | project name, type, id, resourceGroup"
            $relatedResources = Search-AzGraph -Query $query
            
            foreach ($resource in $relatedResources) {
                # Skip resources we've already identified
            if ($resource.id -eq $vm.Id) { continue }
            if ($results.NetworkInterfaces | Where-Object { $_.Id -eq $resource.id }) { continue }
            if ($results.Disks | Where-Object { $_.Id -eq $resource.id }) { continue }
            if ($results.Extensions | Where-Object { $_.Id -eq $resource.id }) { continue }
            if ($results.ManagedIdentity | Where-Object { $_.Id -eq $resource.id }) { continue }
            if ($results.KeyVault | Where-Object { $_.Id -eq $resource.id }) { continue }
            if ($results.DiskEncryptionSet | Where-Object { $_.Id -eq $resource.id }) { continue }
            
            Write-Output "Found Related Resource: $($resource.name) ($($resource.type))"
            $results.RelatedResources += [PSCustomObject]@{
                Name = $resource.name
                Id = $resource.id
                Type = $resource.type
                ResourceGroup = $resource.resourceGroup
            }
        }
    }
    catch {
        Write-Warning "Error querying Resource Graph: $_"
    }
}

# Output results in tabular format
Write-Output "`n============================================================="
Write-Output "           AVD SESSION HOST RESOURCE INVENTORY"
Write-Output "============================================================="
Write-Output "Session Host: $SessionHostName"
Write-Output "Resource Group: $ResourceGroupName"
Write-Output "Subscription: $((Get-AzContext).Subscription.Name)"
Write-Output "Subscription ID: $((Get-AzContext).Subscription.Id)"
Write-Output "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Output "============================================================="
    
    # Output VM information
    Write-Output "`n[SESSION HOST DETAILS]"
    Write-Output "Name: $($results.SessionHost.Name)"
    Write-Output "Size: $($results.SessionHost.Size)"
    Write-Output "OS Type: $($results.SessionHost.OSType)"
    Write-Output "Location: $($results.SessionHost.Location)"
    Write-Output "Provisioning State: $($results.SessionHost.ProvisioningState)"
    Write-Output "Power State: $($results.SessionHost.PowerState)"
    if ($results.SessionHost.LastPowerOnTime) { Write-Output "Last Powered On: $($results.SessionHost.LastPowerOnTime)" }
    if ($results.SessionHost.LastLoggedInTime) { Write-Output "Last Logged In: $($results.SessionHost.LastLoggedInTime)" }
    if ($results.SessionHost.AssignedUser) { Write-Output "Assigned User: $($results.SessionHost.AssignedUser)" }
    Write-Output "Resource ID: $($results.SessionHost.Id)"
    
    # Output Entra ID information
    if ($results.EntraID.DeviceId) {
        Write-Output "`n[ENTRA ID ENROLLMENT]"
        Write-Output "Device ID: $($results.EntraID.DeviceId)"
        Write-Output "Join Type: $($results.EntraID.JoinType)"
        Write-Output "MDM Enrolled: $($results.EntraID.MDMEnrolled)"
        Write-Output "Last Sync: $($results.EntraID.LastSyncTime)"
        Write-Output "Trust Type: $($results.EntraID.TrustType)"
        if ($results.EntraID.EnrollmentDate) { Write-Output "Enrollment Date: $($results.EntraID.EnrollmentDate)" }
    }
    
    # Output Intune information
    if ($results.Intune.Enrolled) {
        Write-Output "`n[INTUNE ENROLLMENT]"
        Write-Output "Device ID: $($results.Intune.DeviceId)"
        Write-Output "Enrolled: $($results.Intune.Enrolled)"
        Write-Output "Enrollment Date: $($results.Intune.EnrollmentDate)"
        Write-Output "Last Contact: $($results.Intune.LastContactTime)"
        Write-Output "Compliance State: $($results.Intune.ComplianceState)"
        Write-Output "Management Agent: $($results.Intune.ManagementAgent)"
    }
    
    # Output Network Interfaces
    if ($results.NetworkInterfaces -and $results.NetworkInterfaces.Count -gt 0) {
        Write-TableOutput -Data $results.NetworkInterfaces -Title "NETWORK INTERFACES" -Properties @("Name", "PrivateIpAddress")
    }
    
    # Output Disks
    if ($results.Disks -and $results.Disks.Count -gt 0) {
        Write-TableOutput -Data $results.Disks -Title "DISKS" -Properties @("Name", "DiskType", "SizeGB", "Sku", "Encryption", "DiskState")
    }
    
    # Output Managed Identity
    if ($results.ManagedIdentity -and $results.ManagedIdentity.Count -gt 0) {
        Write-TableOutput -Data $results.ManagedIdentity -Title "MANAGED IDENTITY" -Properties @("Name", "PrincipalId", "ClientId")
    }
    
    # Output Key Vault
    if ($results.KeyVault -and $results.KeyVault.Count -gt 0) {
        Write-TableOutput -Data $results.KeyVault -Title "KEY VAULT" -Properties @("Name", "VaultUri", "EnabledForDiskEncryption")
    }
    
    # Output Disk Encryption Set
    if ($results.DiskEncryptionSet -and $results.DiskEncryptionSet.Count -gt 0) {
        Write-TableOutput -Data $results.DiskEncryptionSet -Title "DISK ENCRYPTION SET" -Properties @("Name", "EncryptionType", "KeyVaultUrl")
    }
    
    # Output AVD Host Pools
    if ($results.AVD.HostPools -and $results.AVD.HostPools.Count -gt 0) {
        Write-TableOutput -Data $results.AVD.HostPools -Title "AVD HOST POOLS" -Properties @("Name", "Type", "LoadBalancerType")
    }
    
    # Output AVD Session Hosts
    if ($results.AVD.SessionHosts -and $results.AVD.SessionHosts.Count -gt 0) {
        Write-TableOutput -Data $results.AVD.SessionHosts -Title "AVD SESSION HOST REGISTRATIONS" -Properties @("Name", "Status", "AssignedUser", "HostPoolName")
    }
    
    # Output Extensions
    if ($results.Extensions -and $results.Extensions.Count -gt 0) {
        Write-TableOutput -Data $results.Extensions -Title "VM EXTENSIONS" -Properties @("Name", "Publisher", "ExtensionType", "ProvisioningState")
    }
    
    # Output Related Resources
    if ($results.RelatedResources -and $results.RelatedResources.Count -gt 0) {
        Write-TableOutput -Data $results.RelatedResources -Title "OTHER RELATED RESOURCES" -Properties @("Name", "Type", "ResourceGroup")
    }
    
    Write-Output "`n============================================================="
    Write-Output "END OF REPORT"
    Write-Output "============================================================="
    
# Add this right before the decommissioning check
Write-OrderedInventorySummary -InventoryResults $results

# Check if decommissioning is requested
if ($Decommission) {
    # Store Graph token globally if we have one for use in decommissioning
    if ($graphToken) {
        $global:graphToken = $graphToken
    }
    
    # First run in What-If mode to show what would happen
    Write-Output "`nRunning decommission in simulation mode first...`n"
    
    # Actually perform the decommission if confirmed
    if ($ConfirmDecommission) {
        Write-Output "`n!!! PROCEEDING WITH ACTUAL DECOMMISSIONING !!!`n"
        Decommission-AVDSessionHost -InventoryResults $results
    }
    else {
        Write-Output "`nThis is a simulation only. To perform actual deletion, run with -Decommission -ConfirmDecommission"
        Decommission-AVDSessionHost -InventoryResults $results -WhatIf
    }
}



    # Return the complete data object for any further processing
    return $results
}
catch {
    $errorMessage = "Error: $($_.Exception.Message)"
    Write-Error $errorMessage
    
    # Provide more detailed error information if available
    if ($_.Exception.Response) {
        Write-Error "Status code: $($_.Exception.Response.StatusCode.value__)"
        Write-Error "Status description: $($_.Exception.Response.StatusDescription)"
    }
    
    if ($_.ScriptStackTrace) {
        Write-Error "Script stack trace: $($_.ScriptStackTrace)"
    }
    
    throw $_.Exception
}
