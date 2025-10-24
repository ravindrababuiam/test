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
.NOTES
    Required Microsoft Graph API permissions:
    - Device.Read.All: To read device information from Entra ID
    - DeviceManagementManagedDevices.Read.All: To read device information from Intune
    - AuditLog.Read.All: To read sign-in logs for LastSignInDate
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SessionHostName,

    [Parameter(Mandatory = $false)]
    [string]$CMDBLastDiscoveryDate,
    
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
    [bool]$ConfirmDecommission = $false,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAvdRemoval = $false,

    [Parameter(Mandatory = $false)]
    [int]$InactivityThresholdDays = 45


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
   
# 1.5. Application Groups and MA Groups
if ($InventoryResults.AVD.ApplicationGroups -and $InventoryResults.AVD.ApplicationGroups.Count -gt 0) {
    Write-Output "`n[1.5. APPLICATION GROUPS AND MA GROUPS]"
    foreach ($appGroup in $InventoryResults.AVD.ApplicationGroups) {
        Write-Output "Application Group: $($appGroup.Name)"
        Write-Output "Type: $($appGroup.Type)"
        
        # List associated MA groups
        $associatedMAGroups = $InventoryResults.AVD.MAGroups | Where-Object { $_.ApplicationGroup -eq $appGroup.Name }
        if ($associatedMAGroups) {
            foreach ($maGroup in $associatedMAGroups) {
                Write-Output "  - MA Group: $($maGroup.Name)"
            }
        }
        else {
            Write-Output "  - No associated MA groups found"
        }
        Write-Output "------------------------"
    }
} 
else {
    Write-Output "`n[1.5. APPLICATION GROUPS AND MA GROUPS] - None found"
}

# 2. Intune Enrollment
if ($InventoryResults.Intune.DeviceId) {
    Write-Output "`n[2. INTUNE ENROLLMENT]"
    Write-Output "Device ID: $($InventoryResults.Intune.DeviceId)"
    Write-Output "Enrolled: $($InventoryResults.Intune.Enrolled)"
    if ($InventoryResults.Intune.EnrollmentDate) {
        Write-Output "Enrollment Date: $($InventoryResults.Intune.EnrollmentDate)"
    }
    if ($InventoryResults.Intune.LastSignInDate) {
        Write-Output "Last Sign In: $($InventoryResults.Intune.LastSignInDate)"
        if ($InventoryResults.Intune.SignInApplication) {
            Write-Output "Sign-In App: $($InventoryResults.Intune.SignInApplication)"
        }
        if ($InventoryResults.Intune.PSObject.Properties.Name -contains "SignInDevice" -and $InventoryResults.Intune.SignInDevice) {
            Write-Output "Sign-In Device: $($InventoryResults.Intune.SignInDevice)"
        }
        if ($InventoryResults.Intune.PSObject.Properties.Name -contains "SignInNotes" -and $InventoryResults.Intune.SignInNotes) {
            Write-Output "Sign-In Notes: $($InventoryResults.Intune.SignInNotes)"
        }
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
    
# 3.5. CMDB Information
Write-Output "`n[3.5. CMDB INFORMATION]"
if ($null -ne $InventoryResults.CMDBLastDiscovery) {
    Write-Output "Last Discovery: $($InventoryResults.CMDBLastDiscovery)"
    Write-Output "Days Since Discovery: $((New-TimeSpan -Start $InventoryResults.CMDBLastDiscovery -End (Get-Date)).Days)"
    
    # Add CMDB information to decommission recommendation if not already there
    if (-not $InventoryResults.DecommissionRecommendation.Reason.Contains("CMDB")) {
        $InventoryResults.DecommissionRecommendation.Reason += " (CMDB last discovery: $($InventoryResults.CMDBLastDiscovery))"
    }
} else {
    if ($CMDBLastDiscoveryDate) {
        Write-Output "CMDB Discovery Date was provided but could not be parsed: $CMDBLastDiscoveryDate"
    } else {
        Write-Output "No CMDB Discovery Date provided"
    }
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
    
    # 10. Decommission Recommendation
    Write-Output "`n[10. DECOMMISSION RECOMMENDATION]"
    Write-Output "Recommendation: $(if($InventoryResults.DecommissionRecommendation.IsRecommended){'RECOMMENDED FOR DECOMMISSION'}else{'KEEP'})"
    Write-Output "Reason: $($InventoryResults.DecommissionRecommendation.Reason)"
    Write-Output "Last Activity: $($InventoryResults.DecommissionRecommendation.LastActivityDays) days ago"
    Write-Output "Inactivity Threshold: $($InventoryResults.DecommissionRecommendation.InactivityThreshold) days"
    

    Write-Output "`n============================================================="
}

function Analyze-InactivityStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$InventoryResults,
        
        [Parameter(Mandatory = $false)]
        [int]$InactivityThresholdDays = 45
    )
    
    # Initialize with no recommendation
    $InventoryResults.DecommissionRecommendation.InactivityThreshold = $InactivityThresholdDays
    $InventoryResults.DecommissionRecommendation.IsRecommended = $false
    $InventoryResults.DecommissionRecommendation.Reason = "Active"
    
# Establish priority order for timestamp sources

# Debug: Show types of timestamp values
Write-Output "EntraID LastActivity: $($InventoryResults.EntraID.LastActivity) [Type: $(if($InventoryResults.EntraID.LastActivity -ne $null){$InventoryResults.EntraID.LastActivity.GetType().FullName}else{'null'})]"
Write-Output "Intune LastSignInDate: $($InventoryResults.Intune.LastSignInDate) [Type: $(if($InventoryResults.Intune.LastSignInDate -ne $null){$InventoryResults.Intune.LastSignInDate.GetType().FullName}else{'null'})]"
Write-Output "CMDB LastDiscovery: $($InventoryResults.CMDBLastDiscovery) [Type: $(if($InventoryResults.CMDBLastDiscovery -ne $null){$InventoryResults.CMDBLastDiscovery.GetType().FullName}else{'null'})]"
Write-Output "AVD LastHeartbeat: $($InventoryResults.AVD.LastHeartbeat) [Type: $(if($InventoryResults.AVD.LastHeartbeat -ne $null){$InventoryResults.AVD.LastHeartbeat.GetType().FullName}else{'null'})]"
Write-Output "SessionHost LastLoggedInTime: $($InventoryResults.SessionHost.LastLoggedInTime) [Type: $(if($InventoryResults.SessionHost.LastLoggedInTime -ne $null){$InventoryResults.SessionHost.LastLoggedInTime.GetType().FullName}else{'null'})]"
    # Priority 1: Intune/EntraID data (highest priority)
    # Priority 2: CMDB data (enterprise-wide discovery)
    # Priority 3: AVD-specific metrics (lowest priority)
    
    # Initialize collections to store timestamps by priority level
    $priority1Timestamps = @()
    $priority1Reasons = @()
    
    $priority2Timestamps = @()
    $priority2Reasons = @()
    
    $priority3Timestamps = @()
    $priority3Reasons = @()


# Validate all timestamps are proper DateTime objects
Write-Output "Validating timestamps before analysis..."

for ($i = 0; $i -lt $priority1Timestamps.Count; $i++) {
    if (-not ($priority1Timestamps[$i] -is [DateTime])) {
        Write-Warning "Invalid timestamp in Priority 1: '$($priority1Timestamps[$i])' ($($priority1Reasons[$i])) - Not a DateTime object"
        # Attempt to convert string to DateTime if possible
        if ($priority1Timestamps[$i] -is [string] -and ![string]::IsNullOrEmpty($priority1Timestamps[$i])) {
            try {
                $priority1Timestamps[$i] = [DateTime]::Parse($priority1Timestamps[$i])
                Write-Output "  Successfully converted string to DateTime: $($priority1Timestamps[$i])"
            } catch {
                # Remove the invalid timestamp and its reason
                $priority1Timestamps.RemoveAt($i)
                $priority1Reasons.RemoveAt($i)
                $i--  # Adjust index since we removed an item
            }
        } else {
            # Remove the invalid timestamp and its reason
            $priority1Timestamps.RemoveAt($i)
            $priority1Reasons.RemoveAt($i)
            $i--  # Adjust index since we removed an item
        }
    }
}

# Apply the same validation to priority 2 and 3
for ($i = 0; $i -lt $priority2Timestamps.Count; $i++) {
    if (-not ($priority2Timestamps[$i] -is [DateTime])) {
        Write-Warning "Invalid timestamp in Priority 2: '$($priority2Timestamps[$i])' ($($priority2Reasons[$i])) - Not a DateTime object"
        # Attempt to convert string to DateTime if possible
        if ($priority2Timestamps[$i] -is [string] -and ![string]::IsNullOrEmpty($priority2Timestamps[$i])) {
            try {
                $priority2Timestamps[$i] = [DateTime]::Parse($priority2Timestamps[$i])
                Write-Output "  Successfully converted string to DateTime: $($priority2Timestamps[$i])"
            } catch {
                # Remove the invalid timestamp and its reason
                $priority2Timestamps.RemoveAt($i)
                $priority2Reasons.RemoveAt($i)
                $i--  # Adjust index since we removed an item
            }
        } else {
            # Remove the invalid timestamp and its reason
            $priority2Timestamps.RemoveAt($i)
            $priority2Reasons.RemoveAt($i)
            $i--  # Adjust index since we removed an item
        }
    }
}

for ($i = 0; $i -lt $priority3Timestamps.Count; $i++) {
    if (-not ($priority3Timestamps[$i] -is [DateTime])) {
        Write-Warning "Invalid timestamp in Priority 3: '$($priority3Timestamps[$i])' ($($priority3Reasons[$i])) - Not a DateTime object"
        # Attempt to convert string to DateTime if possible
        if ($priority3Timestamps[$i] -is [string] -and ![string]::IsNullOrEmpty($priority3Timestamps[$i])) {
            try {
                $priority3Timestamps[$i] = [DateTime]::Parse($priority3Timestamps[$i])
                Write-Output "  Successfully converted string to DateTime: $($priority3Timestamps[$i])"
            } catch {
                # Remove the invalid timestamp and its reason
                $priority3Timestamps.RemoveAt($i)
                $priority3Reasons.RemoveAt($i)
                $i--  # Adjust index since we removed an item
            }
        } else {
            # Remove the invalid timestamp and its reason
            $priority3Timestamps.RemoveAt($i)
            $priority3Reasons.RemoveAt($i)
            $i--  # Adjust index since we removed an item
        }
    }
}


# PRIORITY 1: Intune/EntraID data (highest priority) - ONLY DEVICE-SPECIFIC SIGNALS
Write-Output "Processing Priority 1 timestamps (device-specific signals)..."

# Add EntraID LastActivity if available and verified as a DateTime
if ($InventoryResults.EntraID.LastActivity -is [DateTime]) {
    $priority1Timestamps += $InventoryResults.EntraID.LastActivity
    $priority1Reasons += "EntraID LastActivity"
    Write-Output "  Added EntraID LastActivity: $($InventoryResults.EntraID.LastActivity)"
} elseif ($InventoryResults.EntraID.LastActivity) {
    Write-Warning "  EntraID LastActivity is not a valid DateTime: $($InventoryResults.EntraID.LastActivity)"
    try {
        $convertedDate = [DateTime]::Parse("$($InventoryResults.EntraID.LastActivity)")
        $priority1Timestamps += $convertedDate
        $priority1Reasons += "EntraID LastActivity (converted)"
        Write-Output "  Added converted EntraID LastActivity: $convertedDate"
    } catch {
        Write-Warning "  Could not convert EntraID LastActivity to DateTime: $($InventoryResults.EntraID.LastActivity)"
    }
}

# Add Intune LastSignInDate if available and ONLY if it's from this specific device
if ($InventoryResults.Intune.LastSignInDate) {
    # Only include the sign-in if it's device-specific
    $isDeviceSpecificSignIn = $false
    $deviceSpecificEvidence = ""
    
    # Check if the sign-in is specifically from this device using multiple criteria
    # 1. Device name match
    if ($InventoryResults.Intune.PSObject.Properties.Name -contains "SignInDevice" -and 
        $InventoryResults.Intune.SignInDevice) {
        
        $vmName = $InventoryResults.SessionHost.Name
        $signInDevice = $InventoryResults.Intune.SignInDevice
        
        # Exact match
        if ($signInDevice -eq $vmName) {
            $isDeviceSpecificSignIn = $true
            $deviceSpecificEvidence = "Exact device name match"
        }
        # VM name contains sign-in device name
        elseif ($vmName -like "*$signInDevice*") {
            $isDeviceSpecificSignIn = $true
            $deviceSpecificEvidence = "VM name contains sign-in device"
        }
        # Sign-in device contains VM name
        elseif ($signInDevice -like "*$vmName*") {
            $isDeviceSpecificSignIn = $true
            $deviceSpecificEvidence = "Sign-in device contains VM name"
        }
    }
    
    # 2. Check explicit flag from audit logs processing
    if (-not $isDeviceSpecificSignIn -and 
        $InventoryResults.Intune.PSObject.Properties.Name -contains "IsDeviceSpecificSignIn" -and 
        $InventoryResults.Intune.IsDeviceSpecificSignIn -eq $true) {
        $isDeviceSpecificSignIn = $true
        $deviceSpecificEvidence = "Marked as device-specific in audit logs"
    }
    
    # Only include the timestamp if it's device-specific
    if ($isDeviceSpecificSignIn) {
        # Ensure it's a DateTime object
        if ($InventoryResults.Intune.LastSignInDate -is [DateTime]) {
            $priority1Timestamps += $InventoryResults.Intune.LastSignInDate
            $priority1Reasons += "User Sign-In on device $($InventoryResults.Intune.SignInDevice) ($deviceSpecificEvidence)"
            Write-Output "  Added Intune LastSignInDate (device-specific): $($InventoryResults.Intune.LastSignInDate)"
        } else {
            try {
                $convertedDate = [DateTime]::Parse("$($InventoryResults.Intune.LastSignInDate)")
                $priority1Timestamps += $convertedDate
                $priority1Reasons += "User Sign-In on device $($InventoryResults.Intune.SignInDevice) ($deviceSpecificEvidence, converted)"
                Write-Output "  Added converted Intune LastSignInDate: $convertedDate"
            } catch {
                Write-Warning "  Could not convert Intune LastSignInDate to DateTime: $($InventoryResults.Intune.LastSignInDate)"
            }
        }
    }
    # Explicitly ignore sign-ins that are not device-specific
    else {
        $reason = "Not matching any device criteria"
        if ($InventoryResults.Intune.PSObject.Properties.Name -contains "SignInNotes") {
            $reason = $InventoryResults.Intune.SignInNotes
        }
        Write-Output "  Ignoring non-specific sign-in: $reason"
    }
}
    
    # PRIORITY 2: CMDB data
    # Add CMDB LastDiscoveryDate if available
    if ($InventoryResults.CMDBLastDiscovery) {
        $priority2Timestamps += $InventoryResults.CMDBLastDiscovery
        $priority2Reasons += "CMDB LastDiscoveryDate"
    }
    
# PRIORITY 3: AVD-specific metrics (lowest priority)
    # Add AVD LastHeartbeat if available
    if ($InventoryResults.AVD.LastHeartbeat) {
        $priority3Timestamps += $InventoryResults.AVD.LastHeartbeat
        $priority3Reasons += "AVD LastHeartbeat"
    }
    
    # Add Session LastLoggedInTime if available
    if ($InventoryResults.SessionHost.LastLoggedInTime) {
        $priority3Timestamps += $InventoryResults.SessionHost.LastLoggedInTime
        $priority3Reasons += "AVD LastLoggedInTime"
    }



# Now check timestamps in priority order
$mostRecentActivity = $null
$mostRecentReason = "No activity data found"
$foundValidTimestamp = $false

# Check Priority 1 first (Intune/EntraID - highest priority)
if ($priority1Timestamps.Count -gt 0) {
    $mostRecentActivity = ($priority1Timestamps | Sort-Object -Descending)[0]
    $mostRecentReason = $priority1Reasons[$priority1Timestamps.IndexOf($mostRecentActivity)]
    $foundValidTimestamp = $true
    Write-Output "Found Priority 1 (Intune/EntraID) timestamp: $mostRecentActivity - $mostRecentReason"
}
    # If no Priority 1, check Priority 2 (CMDB)
    elseif ($priority2Timestamps.Count -gt 0) {
        $mostRecentActivity = ($priority2Timestamps | Sort-Object -Descending)[0]
        $mostRecentReason = $priority2Reasons[$priority2Timestamps.IndexOf($mostRecentActivity)]
        $foundValidTimestamp = $true
        Write-Output "Found Priority 2 (CMDB) timestamp: $mostRecentActivity - $mostRecentReason"
    }
# If no Priority 1 or 2, check Priority 3 (AVD - lowest priority)
elseif ($priority3Timestamps.Count -gt 0) {
    $mostRecentActivity = ($priority3Timestamps | Sort-Object -Descending)[0]
    $mostRecentReason = $priority3Reasons[$priority3Timestamps.IndexOf($mostRecentActivity)]
    $foundValidTimestamp = $true
    Write-Output "Found Priority 3 (AVD) timestamp: $mostRecentActivity - $mostRecentReason"
}

# Calculate days since last activity if we found a valid timestamp
if ($foundValidTimestamp) {
    # Ensure mostRecentActivity is a DateTime object
    if (-not ($mostRecentActivity -is [DateTime])) {
        Write-Warning "Most recent activity is not a valid DateTime object. Attempting to convert."
        try {
            $mostRecentActivity = [DateTime]::Parse("$mostRecentActivity")
            Write-Output "Successfully converted most recent activity to DateTime: $mostRecentActivity"
        } catch {
            Write-Warning "Failed to convert most recent activity to DateTime. Setting to minimum date."
            $mostRecentActivity = [DateTime]::MinValue
        }
    }
    
    # Additional safety check before calculation
    if ($mostRecentActivity -ne [DateTime]::MinValue) {
        $currentDate = Get-Date
        Write-Output "Calculating days between $mostRecentActivity and $currentDate"
        $timeSpan = New-TimeSpan -Start $mostRecentActivity -End $currentDate
        $daysSinceActivity = [Math]::Round($timeSpan.TotalDays)
        
        $InventoryResults.DecommissionRecommendation.LastActivityDays = $daysSinceActivity
        Write-Output "Days since activity: $daysSinceActivity (from $mostRecentReason)"
        
        # Check if machine has been inactive for longer than threshold
        if ($daysSinceActivity -gt $InactivityThresholdDays) {
            $InventoryResults.DecommissionRecommendation.IsRecommended = $true
            $InventoryResults.DecommissionRecommendation.Reason = "Inactive for $daysSinceActivity days (last activity: $mostRecentReason on $mostRecentActivity)"
        }
        else {
            $InventoryResults.DecommissionRecommendation.Reason = "Active within $daysSinceActivity days (last activity: $mostRecentReason on $mostRecentActivity)"
        }
    } else {
        Write-Warning "Cannot calculate days since activity - invalid date."
        $InventoryResults.DecommissionRecommendation.LastActivityDays = 999 # Clear indication of problem
        $InventoryResults.DecommissionRecommendation.Reason = "Unable to determine activity status - invalid date format"
    }
}
else {
    Write-Warning "No valid timestamps found for activity calculation"
    $InventoryResults.DecommissionRecommendation.LastActivityDays = 999 # Clear indication of problem
    $InventoryResults.DecommissionRecommendation.Reason = "Unable to determine activity status - no timestamps available"
}
    
    return $InventoryResults
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
    Write-Output "Decommission Recommendation: RECOMMENDED FOR DECOMMISSION"
    Write-Output "Reason: $($InventoryResults.DecommissionRecommendation.Reason)"
    Write-Output "Last Activity: $($InventoryResults.DecommissionRecommendation.LastActivityDays) days ago"
    Write-Output "Inactivity Threshold: $($InventoryResults.DecommissionRecommendation.InactivityThreshold) days"
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
    # STEP 3.5: Remove User from MA Group
if ($InventoryResults.AVD.SessionHosts -and $InventoryResults.AVD.SessionHosts[0].AssignedUser) {
    Write-Output "`nSTEP 3.5: Removing user from MA Group..."
    $assignedUser = $InventoryResults.AVD.SessionHosts[0].AssignedUser
    $hostPoolName = $InventoryResults.AVD.SessionHosts[0].HostPoolName
    $hostPoolRg = ($InventoryResults.AVD.HostPools | Where-Object { $_.Name -eq $hostPoolName }).ResourceGroup
    
    try {
        if (-not $WhatIf) {
            # Using Microsoft Graph REST API with the existing token
            $graphToken = $global:graphToken
            
            if ($graphToken) {
                $graphHeaders = @{
                    "Authorization" = "Bearer $graphToken"
                    "Content-Type" = "application/json"
                }
                
                # Find the application group associated with the host pool
                Write-Output "Finding application group for host pool: $hostPoolName"
                $appGroups = Get-AzWvdApplicationGroup -ResourceGroupName $hostPoolRg
                $appGroup = $appGroups | Where-Object { $_.HostPoolArmPath -like "*$hostPoolName*" } | Select-Object -First 1
                
                if ($appGroup) {
                    # Find MA group in role assignments
                    Write-Output "Looking for MA group in application group: $($appGroup.Name)"
                    $roleAssignments = Get-AzRoleAssignment -Scope $appGroup.Id
                    $maGroup = $null
                    
                    foreach ($assignment in $roleAssignments) {
                        if ($assignment.ObjectType -eq "Group" -and 
                            $assignment.DisplayName -like "MA*" -and 
                            $assignment.DisplayName -like "*-AG-W11-*" -and 
                            ($assignment.DisplayName -like "*STD*" -or $assignment.DisplayName -like "*DEV*")) {
                            
                            $maGroup = $assignment
                            Write-Output "Found matching MA group: $($maGroup.DisplayName)"
                            break
                        }
                    }
                    
                    if ($maGroup) {
                        # Find user's objectId
                        Write-Output "Getting objectId for user: $assignedUser"
                        $userFilter = [System.Web.HttpUtility]::UrlEncode("userPrincipalName eq '$assignedUser'")
                        $userUrl = "https://graph.microsoft.com/v1.0/users?`$filter=$userFilter"
                        
                        $userResponse = Invoke-RestMethod -Uri $userUrl -Headers $graphHeaders -Method GET -ErrorAction Stop
                        
                        if ($userResponse.value -and $userResponse.value.Count -gt 0) {
                            $userId = $userResponse.value[0].id
                            
                            # Remove user from MA group
                            Write-Output "Removing user $assignedUser ($userId) from MA group $($maGroup.DisplayName)..."
                            $removeUrl = "https://graph.microsoft.com/v1.0/groups/$($maGroup.ObjectId)/members/$userId/`$ref"
                            Invoke-RestMethod -Uri $removeUrl -Headers $graphHeaders -Method DELETE -ErrorAction Stop
                            
                            Write-DecommStatus -Resource "MA Group" -Status "Success" -Message "Removed user $assignedUser from MA group $($maGroup.DisplayName)"
                        }
                        else {
                            Write-DecommStatus -Resource "MA Group" -Status "Failed" -Message "Could not find user $assignedUser in Azure AD"
                        }
                    }
                    else {
                        Write-DecommStatus -Resource "MA Group" -Status "Skipped" -Message "No matching MA group found for app group $($appGroup.Name)"
                    }
                }
                else {
                    Write-DecommStatus -Resource "MA Group" -Status "Skipped" -Message "No application group found for host pool $hostPoolName"
                }
            }
            else {
                Write-DecommStatus -Resource "MA Group" -Status "Skipped" -Message "No Graph API token available, manual removal required for user: $assignedUser"
            }
        }
        else {
            Write-DecommStatus -Resource "MA Group" -Status "Skipped" -Message "Would remove user $assignedUser from associated MA group"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DecommStatus -Resource "MA Group" -Status "Failed" -Message "Could not remove user from MA group - $errorMsg"
    }
}
else {
    Write-Output "`nSTEP 3.5: No user assigned to this session host, skipping MA group removal"
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
        ApplicationGroups = @()
        MAGroups = @()
        LastHeartbeat = $null         # Add this line
    }
    EntraID = [PSCustomObject]@{
        DeviceId = $null
        JoinType = $null
        MDMEnrolled = $null
        LastSyncTime = $null
        LastActivity = $null          # Add this line
        TrustType = $null
        EnrollmentDate = $null
    }

    Intune = [PSCustomObject]@{
        DeviceId = $null
        Enrolled = $false
        EnrollmentDate = $null
        LastContactTime = $null
        LastSignInDate = $null
        ComplianceState = $null
        ManagementAgent = $null
        SignInApplication = $null
        SignInDevice = $null
        SignInLocation = $null
        SignInNotes = $null
    }


    DecommissionRecommendation = [PSCustomObject]@{  
        IsRecommended = $false
        Reason = ""
        LastActivityDays = 0
        InactivityThreshold = $InactivityThresholdDays
        }
    RelatedResources = @()
}
# Parse CMDB date first into a variable
$parsedCMDBDate = $null
if ($CMDBLastDiscoveryDate) {
    Write-Output "Attempting to parse CMDB date: $CMDBLastDiscoveryDate"
    try {
        $parsedCMDBDate = [DateTime]::ParseExact($CMDBLastDiscoveryDate, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Output "Successfully parsed CMDB date: $parsedCMDBDate"
    } catch {
        Write-Warning "Unable to parse CMDB date with exact format. Trying general parsing..."
        try {
            $parsedCMDBDate = [DateTime]::Parse($CMDBLastDiscoveryDate)
            Write-Output "Successfully parsed CMDB date with general parsing: $parsedCMDBDate"
        } catch {
            Write-Warning "Failed to parse CMDB date: $CMDBLastDiscoveryDate. Error: $($_.Exception.Message)"
        }
    }
} else {
    Write-Output "No CMDB date provided."
}

# Add the parsed date to results
$results | Add-Member -NotePropertyName "CMDBLastDiscovery" -NotePropertyValue $parsedCMDBDate
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
            $global:graphToken = $graphToken  # Store token globally for later use
            
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
# Try to ensure we have a valid DateTime object for LastSyncTime
if ($device.approximateLastSignInDateTime -is [DateTime]) {
    $results.EntraID.LastSyncTime = $device.approximateLastSignInDateTime
}
elseif ($device.approximateLastSignInDateTime -is [string] -and ![string]::IsNullOrEmpty($device.approximateLastSignInDateTime)) {
    try {
        $results.EntraID.LastSyncTime = [DateTime]::Parse($device.approximateLastSignInDateTime)
    }
    catch {
        Write-Warning "  Could not parse EntraID LastSyncTime as DateTime: $($device.approximateLastSignInDateTime)"
        $results.EntraID.LastSyncTime = $null
    }
}
else {
    $results.EntraID.LastSyncTime = $null
}
# Try to ensure we have a valid DateTime object for LastActivity
if ($device.approximateLastSignInDateTime -is [DateTime]) {
    $results.EntraID.LastActivity = $device.approximateLastSignInDateTime
    Write-Output "  EntraID LastActivity is a valid DateTime: $($results.EntraID.LastActivity)"
}
elseif ($device.approximateLastSignInDateTime -is [string] -and ![string]::IsNullOrEmpty($device.approximateLastSignInDateTime)) {
    try {
        $results.EntraID.LastActivity = [DateTime]::Parse($device.approximateLastSignInDateTime)
        Write-Output "  Converted EntraID LastActivity string to DateTime: $($results.EntraID.LastActivity)"
    }
    catch {
        Write-Warning "  Could not parse EntraID LastActivity as DateTime: $($device.approximateLastSignInDateTime)"
        $results.EntraID.LastActivity = $null
    }
}
else {
    Write-Warning "  EntraID LastActivity is not a valid DateTime: $($device.approximateLastSignInDateTime)"
    $results.EntraID.LastActivity = $null
}
                    $results.EntraID.TrustType = $device.trustType
# Ensure EnrollmentDate is a valid DateTime
if ($device.registrationDateTime -is [DateTime]) {
    $results.EntraID.EnrollmentDate = $device.registrationDateTime
}
elseif ($device.registrationDateTime -is [string] -and ![string]::IsNullOrEmpty($device.registrationDateTime)) {
    try {
        $results.EntraID.EnrollmentDate = [DateTime]::Parse($device.registrationDateTime)
    }
    catch {
        Write-Warning "  Could not parse EntraID EnrollmentDate as DateTime: $($device.registrationDateTime)"
        $results.EntraID.EnrollmentDate = $null
    }
}
else {
    $results.EntraID.EnrollmentDate = $null
}
                    
                    Write-Output "Found Entra ID device: $($device.id)"
                    Write-Output "  Join Type: $($device.joinType)"
                    Write-Output "  MDM Enrolled: $($results.EntraID.MDMEnrolled)"
                    Write-Output "  Last Sync Time: $($device.approximateLastSignInDateTime)"
                    Write-Output "  Last Sync Time: $(if($results.EntraID.LastSyncTime){$results.EntraID.LastSyncTime}else{'Not available'})"
                    Write-Output "  Last Activity: $($device.approximateLastSignInDateTime)"  # Add this line
                    Write-Output "  Last Activity: $(if($results.EntraID.LastActivity){$results.EntraID.LastActivity}else{'Not available'})"
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
                        $results.Intune.LastSignInDate = $intuneDevice.lastSignInDateTime  # Add this line
                        $results.Intune.ComplianceState = $intuneDevice.complianceState
                        $results.Intune.ManagementAgent = $intuneDevice.managementAgent
                        
                        Write-Output "Found Intune device: $($intuneDevice.id)"
                        Write-Output "  Enrollment Date: $($intuneDevice.enrolledDateTime)"
                        Write-Output "  Last Contact Time: $($intuneDevice.lastSyncDateTime)"
                        Write-Output "  Last Sign In Date: $($intuneDevice.lastSignInDateTime)"  # Add this line
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
                            LastHeartbeat = $sessionHost.LastHeartbeat  # Add this line
                        }
                        
                        # Store the most recent heartbeat in the main AVD results
                        if ($sessionHost.LastHeartbeat) {
                            $results.AVD.LastHeartbeat = $sessionHost.LastHeartbeat
                            Write-Output "  Last Heartbeat: $($sessionHost.LastHeartbeat)"
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
                        
                        # After finding the host pool containing our session host, find the application groups
                        Write-Output "Finding application groups for host pool: $($hostPool.Name)"
                        $hostPoolResourceGroup = $hostPool.Id.Split('/')[4]
                        $appGroups = Get-AzWvdApplicationGroup -ResourceGroupName $hostPoolResourceGroup
                        $associatedAppGroups = $appGroups | Where-Object { $_.HostPoolArmPath -eq $hostPool.Id }
                        
                        foreach ($appGroup in $associatedAppGroups) {
                            Write-Output "Found application group: $($appGroup.Name) for host pool: $($hostPool.Name)"
                            $results.AVD.ApplicationGroups += [PSCustomObject]@{
                                Name = $appGroup.Name
                                Id = $appGroup.Id
                                ResourceGroup = $appGroup.ResourceGroupName
                                Type = $appGroup.ApplicationGroupType
                                HostPoolName = $hostPool.Name
                            }
                            
                            # Find MA groups in role assignments
                            Write-Output "Looking for MA groups in application group: $($appGroup.Name)"
                            $roleAssignments = Get-AzRoleAssignment -Scope $appGroup.Id
                            
                            foreach ($assignment in $roleAssignments) {
                                if ($assignment.ObjectType -eq "Group" -and 
                                    $assignment.DisplayName -like "MA*" -and 
                                    $assignment.DisplayName -like "*-AG-W11-*" -and 
                                    ($assignment.DisplayName -like "*STD*" -or $assignment.DisplayName -like "*DEV*")) {
                                    
                                    Write-Output "Found MA group: $($assignment.DisplayName) for app group: $($appGroup.Name)"
                                    $results.AVD.MAGroups += [PSCustomObject]@{
                                        Name = $assignment.DisplayName
                                        ObjectId = $assignment.ObjectId
                                        ApplicationGroup = $appGroup.Name
                                        HostPoolName = $hostPool.Name
                                    }
                                }
                            }
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
        Write-Output "Last Activity: $($results.EntraID.LastActivity)" 
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
    Write-Output "Last Sign In: $($results.Intune.LastSignInDate)"
    if ($results.Intune.SignInApplication) {
        Write-Output "Sign-In Application: $($results.Intune.SignInApplication)"
    }
    if ($results.Intune.PSObject.Properties.Name -contains "SignInDevice" -and $results.Intune.SignInDevice) {
        Write-Output "Sign-In Device: $($results.Intune.SignInDevice)"
    }
    if ($results.Intune.SignInLocation) {
        Write-Output "Sign-In Location: $($results.Intune.SignInLocation)"
    }
    if ($results.Intune.PSObject.Properties.Name -contains "SignInNotes" -and $results.Intune.SignInNotes) {
        Write-Output "Sign-In Notes: $($results.Intune.SignInNotes)"
    }
    Write-Output "Compliance State: $($results.Intune.ComplianceState)"
    Write-Output "Management Agent: $($results.Intune.ManagementAgent)"
}

# Output CMDB information - enhance this section
if ($results.CMDBLastDiscovery) {
    Write-Output "`n[CMDB INFORMATION]"
    Write-Output "Last Discovery Date: $($results.CMDBLastDiscovery)"
    Write-Output "Days Since Discovery: $((New-TimeSpan -Start $results.CMDBLastDiscovery -End (Get-Date)).Days)"
    Write-Output "Discovery Date Source: Parameter CMDBLastDiscoveryDate"
} else {
    Write-Output "`n[CMDB INFORMATION]"
    if ($CMDBLastDiscoveryDate) {
        Write-Output "CMDB Discovery Date was provided but could not be parsed: $CMDBLastDiscoveryDate"
        Write-Output "Expected format: yyyy-MM-dd HH:mm:ss (e.g., 2025-03-27 13:45:53)"
    } else {
        Write-Output "No CMDB Discovery Date provided"
    }
}

# Create CMDB table for better visualization
if ($results.CMDBLastDiscovery) {
    $cmdbData = @(
        [PSCustomObject]@{
            LastDiscoveryDate = $results.CMDBLastDiscovery
            DaysSinceDiscovery = (New-TimeSpan -Start $results.CMDBLastDiscovery -End (Get-Date)).Days
            Source = "CMDB System"
            
        }
    )
    
    Write-TableOutput -Data $cmdbData -Title "CMDB DISCOVERY DATA" -Properties @("LastDiscoveryDate", "DaysSinceDiscovery")
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
    Write-TableOutput -Data $results.AVD.SessionHosts -Title "AVD SESSION HOST REGISTRATIONS" -Properties @("Name", "Status", "AssignedUser", "HostPoolName", "LastHeartbeat")
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

# Get user sign-in information from audit logs
Write-Output "`nChecking for user sign-in information from audit logs..."

# Make sure we have the Graph token available
if ($global:graphToken) {
    $userUpn = $null
    
   
# Debug: Check if we have assigned user information
Write-Output "DEBUG: Session Host Assigned User: $($results.SessionHost.AssignedUser)"
if ($results.AVD.SessionHosts -and $results.AVD.SessionHosts.Count -gt 0) {
    Write-Output "DEBUG: AVD Session Hosts Assigned Users:"
    foreach ($sessionHost in $results.AVD.SessionHosts) {
        Write-Output "  Host: $($sessionHost.Name), Assigned User: $($sessionHost.AssignedUser)"
    }
}

    # First check if we have an assigned user in the session host
    if ($results.SessionHost.AssignedUser) {
        $userUpn = $results.SessionHost.AssignedUser
        Write-Output "Found assigned user in session host: $userUpn"
    }
    
# If no assigned user, see if we can find it from AVD SessionHosts
elseif ($results.AVD.SessionHosts -and $results.AVD.SessionHosts.Count -gt 0) {
    foreach ($sessionHost in $results.AVD.SessionHosts) {
        if ($sessionHost.AssignedUser) {
            $userUpn = $sessionHost.AssignedUser
            Write-Output "Found assigned user from AVD host: $userUpn"
            break
        }
    }
}

    # If we have a user UPN, get their sign-in logs
    if ($userUpn) {
        # Set up headers for Graph API calls
        $graphHeaders = @{
            "Authorization" = "Bearer $global:graphToken"
            "Content-Type"  = "application/json"
        }
        
        # Encode the UPN for use in URL
        $encodedUpn = [System.Web.HttpUtility]::UrlEncode($userUpn)
        

        # Get the device ID and name for matching in sign-in records
        $deviceId = $null
        if ($results.EntraID.DeviceId) {
            $deviceId = $results.EntraID.DeviceId
        }
        
        # Use the session host name for matching
        $deviceName = $SessionHostName.ToLower()
        
        Write-Output "Querying audit logs for user sign-ins for: $userUpn on device: $deviceName"
        
        # Build the query to get all recent sign-ins for this user (increased to 50 for better chance of finding device-specific sign-ins)
        $signInLogsUrl = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userPrincipalName eq '$encodedUpn'&`$top=50&`$orderby=createdDateTime desc"
        
        try {
            $signInLogsResponse = Invoke-RestMethod -Uri $signInLogsUrl -Headers $graphHeaders -Method GET -ErrorAction Stop
            
            if ($signInLogsResponse.value -and $signInLogsResponse.value.Count -gt 0) {
                Write-Output "Found $($signInLogsResponse.value.Count) sign-in records for user"
                

# First try to find sign-ins specific to this device (exact match)
$deviceSpecificSignIn = $null
$deviceMatchReason = ""

# If we have an Azure AD device ID, try to match on that - most reliable
if ($deviceId) {
    Write-Output "Searching for sign-ins with exact device ID match: $deviceId"
    $deviceSpecificSignIn = $signInLogsResponse.value | 
        Where-Object { $_.deviceDetail.deviceId -eq $deviceId -and $_.status.errorCode -eq 0 } | 
        Sort-Object -Property createdDateTime -Descending | 
        Select-Object -First 1
        
    if ($deviceSpecificSignIn) {
        Write-Output "Found sign-in matching device ID: $deviceId"
        $deviceMatchReason = "Exact device ID match"
    }
}

# If no match by ID, try by device name with exact match first
if (-not $deviceSpecificSignIn) {
    Write-Output "Searching for sign-ins with exact device name match: $deviceName"
    # First try exact match - highest confidence
    $exactNameMatch = $signInLogsResponse.value | 
        Where-Object { 
            $_.deviceDetail.displayName -eq $deviceName -and $_.status.errorCode -eq 0 
        } | 
        Sort-Object -Property createdDateTime -Descending | 
        Select-Object -First 1
        
    if ($exactNameMatch) {
        $deviceSpecificSignIn = $exactNameMatch
        Write-Output "Found sign-in with exact device name match: $deviceName"
        $deviceMatchReason = "Exact device name match"
    }
    # If no exact match, try partial matches - lower confidence
    else {
        Write-Output "No exact name match, trying partial device name matches"
        # VM name contains in sign-in device name or vice versa
        $partialMatch = $signInLogsResponse.value | 
            Where-Object { 
                ($_.deviceDetail.displayName -like "*$deviceName*" -or 
                 ($deviceName -and $deviceName -like "*$($_.deviceDetail.displayName)*")) -and 
                $_.status.errorCode -eq 0 
            } | 
            Sort-Object -Property createdDateTime -Descending | 
            Select-Object -First 1
            
        if ($partialMatch) {
            $deviceSpecificSignIn = $partialMatch
            Write-Output "Found sign-in with partial device name match: $($partialMatch.deviceDetail.displayName)"
            $deviceMatchReason = "Partial device name match"
        }
    }
}

# Also give preference to Windows Sign In app specifically
if (-not $deviceSpecificSignIn) {
    Write-Output "No device match found, looking for 'Windows Sign In' app with Windows OS"
    $windowsSignIn = $signInLogsResponse.value | 
        Where-Object { 
            $_.appDisplayName -eq "Windows Sign In" -and 
            $_.deviceDetail.operatingSystem -eq "Windows" -and
            $_.status.errorCode -eq 0 
        } | 
        Sort-Object -Property createdDateTime -Descending | 
        Select-Object -First 1
        
    if ($windowsSignIn) {
        $deviceSpecificSignIn = $windowsSignIn
        Write-Output "Found Windows Sign In, but device could not be confirmed as: $deviceName"
        $deviceMatchReason = "Windows Sign In (device not confirmed)"
    }
}
                    
                    if ($deviceSpecificSignIn) {
                        Write-Output "Found sign-in matching device name: $deviceName"
                    }
                }
                
# If we found a device-specific sign-in, use that
if ($deviceSpecificSignIn) {
    $results.Intune.LastSignInDate = $deviceSpecificSignIn.createdDateTime
    Write-Output "Last Sign In Date for this device (from audit logs): $($deviceSpecificSignIn.createdDateTime)"
    Write-Output "Sign-in Application: $($deviceSpecificSignIn.appDisplayName)"
    Write-Output "Sign-in Device: $($deviceSpecificSignIn.deviceDetail.displayName)"
    Write-Output "Sign-in Location: $($deviceSpecificSignIn.location.city), $($deviceSpecificSignIn.location.countryOrRegion)"
    
    # Store additional sign-in information
# Store additional sign-in information
$results.Intune | Add-Member -NotePropertyName "SignInApplication" -NotePropertyValue $deviceSpecificSignIn.appDisplayName -Force
$results.Intune | Add-Member -NotePropertyName "SignInDevice" -NotePropertyValue $deviceSpecificSignIn.deviceDetail.displayName -Force
$results.Intune | Add-Member -NotePropertyName "SignInLocation" -NotePropertyValue "$($deviceSpecificSignIn.location.city), $($deviceSpecificSignIn.location.countryOrRegion)" -Force

# Store device match confidence and more detailed info
$isConfirmedDeviceSpecific = $deviceMatchReason -in @("Exact device ID match", "Exact device name match")
$results.Intune | Add-Member -NotePropertyName "SignInMatchReason" -NotePropertyValue $deviceMatchReason -Force

# Add flag to indicate this is a device-specific sign-in - only true for high confidence matches
$results.Intune | Add-Member -NotePropertyName "IsDeviceSpecificSignIn" -NotePropertyValue $isConfirmedDeviceSpecific -Force

# Store match confidence level
$confidenceLevel = switch ($deviceMatchReason) {
    "Exact device ID match" { "High" }
    "Exact device name match" { "High" }
    "Partial device name match" { "Medium" }
    "Windows Sign In (device not confirmed)" { "Low" }
    default { "Unknown" }
}
$results.Intune | Add-Member -NotePropertyName "SignInConfidence" -NotePropertyValue $confidenceLevel -Force

# Add a clear note explaining the match
$noteText = switch ($confidenceLevel) {
    "High" { "This sign-in is confirmed to be from this device ($deviceMatchReason)" }
    "Medium" { "This sign-in is likely from this device but not confirmed ($deviceMatchReason)" }
    "Low" { "This sign-in may not be from this device - using as fallback only ($deviceMatchReason)" }
    default { "Cannot determine if this sign-in is from this device" }
}
$results.Intune | Add-Member -NotePropertyName "SignInNotes" -NotePropertyValue $noteText -Force

Write-Output "Sign-in confidence: $confidenceLevel - $noteText"
}
                else {
                    # For Windows Sign In, match on operatingSystem = Windows
                    $windowsSignIns = $signInLogsResponse.value | 
                        Where-Object { 
                            $_.deviceDetail.operatingSystem -eq "Windows" -and
                            $_.appDisplayName -eq "Windows Sign In" -and
                            $_.status.errorCode -eq 0 
                        } | 
                        Sort-Object -Property createdDateTime -Descending
                    
                    if ($windowsSignIns -and $windowsSignIns.Count -gt 0) {
                        $latestWindowsSignIn = $windowsSignIns[0]
                        
                        # Log all the Windows sign-ins we found for debugging
                        Write-Output "Found $($windowsSignIns.Count) Windows Sign In entries:"
                        foreach ($signin in $windowsSignIns | Select-Object -First 5) {
                            Write-Output "  Time: $($signin.createdDateTime), Device: $($signin.deviceDetail.displayName), OS: $($signin.deviceDetail.operatingSystem)"
                        }
                        
                        
 # Store the Windows sign-in for reference, but mark it as NOT device-specific
$results.Intune.LastSignInDate = $latestWindowsSignIn.createdDateTime
Write-Output "WARNING: Using latest Windows Sign In as fallback (NOT for this specific device)"
Write-Output "Last Windows Sign In (from audit logs): $($latestWindowsSignIn.createdDateTime)"
Write-Output "Sign-in Application: $($latestWindowsSignIn.appDisplayName)"
if ($latestWindowsSignIn.deviceDetail.displayName) {
    Write-Output "Sign-in Device: $($latestWindowsSignIn.deviceDetail.displayName)"
}
Write-Output "Sign-in Location: $($latestWindowsSignIn.location.city), $($latestWindowsSignIn.location.countryOrRegion)"

# Store additional sign-in information but make it clear it's not device-specific
$results.Intune | Add-Member -NotePropertyName "SignInApplication" -NotePropertyValue $latestWindowsSignIn.appDisplayName -Force
if ($latestWindowsSignIn.deviceDetail.displayName) {
    $results.Intune | Add-Member -NotePropertyName "SignInDevice" -NotePropertyValue $latestWindowsSignIn.deviceDetail.displayName -Force
}
$results.Intune | Add-Member -NotePropertyName "SignInLocation" -NotePropertyValue "$($latestWindowsSignIn.location.city), $($latestWindowsSignIn.location.countryOrRegion)" -Force
$results.Intune | Add-Member -NotePropertyName "SignInNotes" -NotePropertyValue "Using latest Windows Sign In (NOT specific to this device)" -Force
$results.Intune | Add-Member -NotePropertyName "IsDeviceSpecificSignIn" -NotePropertyValue $false -Force
                    }
                    else {
                        # If no Windows sign-ins, fall back to last successful sign-in of any type as a last resort
                        $successfulSignIn = $signInLogsResponse.value | Where-Object { $_.status.errorCode -eq 0 } | Select-Object -First 1
                        
                        if ($successfulSignIn) {
                            $results.Intune.LastSignInDate = $successfulSignIn.createdDateTime
                            Write-Output "WARNING: Using latest sign-in of any type as fallback (unlikely to be for this device)"
                            Write-Output "Last Sign In Date (from audit logs): $($successfulSignIn.createdDateTime)"
                            Write-Output "Sign-in Application: $($successfulSignIn.appDisplayName)"
                            if ($successfulSignIn.deviceDetail.displayName) {
                                Write-Output "Sign-in Device: $($successfulSignIn.deviceDetail.displayName)"
                            }
                            Write-Output "Sign-in Location: $($successfulSignIn.location.city), $($successfulSignIn.location.countryOrRegion)"
                            
                            # Store additional sign-in information with warning
                            $results.Intune | Add-Member -NotePropertyName "SignInApplication" -NotePropertyValue $successfulSignIn.appDisplayName -Force
                            if ($successfulSignIn.deviceDetail.displayName) {
                                $results.Intune | Add-Member -NotePropertyName "SignInDevice" -NotePropertyValue $successfulSignIn.deviceDetail.displayName -Force
                            }
                            $results.Intune | Add-Member -NotePropertyName "SignInLocation" -NotePropertyValue "$($successfulSignIn.location.city), $($successfulSignIn.location.countryOrRegion)" -Force
                            $results.Intune | Add-Member -NotePropertyName "SignInNotes" -NotePropertyValue "WARNING: Based on latest sign-in of any type (NOT for this device)" -Force
                            $results.Intune | Add-Member -NotePropertyName "IsDeviceSpecificSignIn" -NotePropertyValue $false -Force
                        }
                        else {
                            Write-Output "No successful sign-ins found in audit logs"
                        }
                    }
                }
            
            else {
                Write-Output "No sign-in records found in audit logs for user $userUpn"
            }
        }
        catch {
            Write-Warning "Error retrieving sign-in logs: $_"
            if ($_.Exception.Response) {
                Write-Warning "Status code: $($_.Exception.Response.StatusCode.value__)"
                Write-Warning "Status description: $($_.Exception.Response.StatusDescription)"
            }
            
            # Check if this might be a permissions issue
            if ($_.Exception.Message -like "*Authorization_RequestDenied*" -or 
                $_.Exception.Message -like "*AccessDenied*" -or 
                $_.Exception.Message -like "*Forbidden*") {
                Write-Warning "This appears to be a permissions issue. Make sure the service principal has the 'AuditLog.Read.All' permission."
                Write-Warning "You might need to add this permission in the Azure Portal under App Registrations > API Permissions."
            }
        }
    }
    else {
        Write-Output "No assigned user found to check for sign-in activity"
    }
    
    # If LastSignInDate is still empty, use LastContactTime as fallback
    if ([string]::IsNullOrEmpty($results.Intune.LastSignInDate) -and $results.Intune.LastContactTime) {
        Write-Output "Using LastContactTime as fallback for LastSignInDate"
        $results.Intune.LastSignInDate = $results.Intune.LastContactTime
    }
} else {
    Write-Warning "No Graph token available to query audit logs"
}

# Add CMDB debug output
Write-Output "CMDB Last Discovery Date Parameter: $CMDBLastDiscoveryDate"
Write-Output "CMDB Last Discovery Parsed: $($results.CMDBLastDiscovery)"
if ($results.CMDBLastDiscovery) {
    Write-Output "CMDB Last Discovery Type: $($results.CMDBLastDiscovery.GetType().FullName)"
}

# Run inactivity analysis to determine if session host should be decommissioned
$results = Analyze-InactivityStatus -InventoryResults $results -InactivityThresholdDays $InactivityThresholdDays

# Add this right before the decommissioning check
Write-OrderedInventorySummary -InventoryResults $results
    


# Check if decommissioning is requested
if ($Decommission) {
    # Store Graph token globally if we have one for use in decommissioning
    if ($graphToken) {
        $global:graphToken = $graphToken
    }
    
    # Check the recommendation before proceeding
    if (-not $results.DecommissionRecommendation.IsRecommended) {
        # If NOT recommended for decommissioning, fail the job
        Write-Output "`n============================================================="
        Write-Output "           DECOMMISSIONING PROCESS HALTED                     "
        Write-Output "============================================================="
        Write-Output "Session Host: $SessionHostName"
        Write-Output "Recommendation: KEEP"
        Write-Output "Reason: $($results.DecommissionRecommendation.Reason)"
        Write-Output "Last Activity: $($results.DecommissionRecommendation.LastActivityDays) days ago"
        Write-Output "Inactivity Threshold: $($results.DecommissionRecommendation.InactivityThreshold) days"
        Write-Output "============================================================="
        
        # Return detailed information for logging
        Write-Error "Decommissioning halted - Session host is NOT recommended for decommissioning."
        Write-Error "Reason: $($results.DecommissionRecommendation.Reason)"
        Write-Error "Last Activity: $($results.DecommissionRecommendation.LastActivityDays) days ago"
        Write-Error "Inactivity Threshold: $($results.DecommissionRecommendation.InactivityThreshold) days"
        
        # Return results but exit with error to fail the job
        return $results
        throw "Decommissioning halted - Session host is NOT recommended for decommissioning."
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
