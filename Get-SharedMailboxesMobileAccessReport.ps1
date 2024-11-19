<#
.SYNOPSIS
    Exchange Server 2019 Shared Mailboxes Mobile Access Report

.DESCRIPTION
    This PowerShell script retrieves all shared mailboxes in an Exchange Server 2019 environment, checks if ActiveSync is enabled for each mailbox, identifies any mobile devices accessing them, gathers detailed statistics for each device, and exports the information into two separate CSV files:
    1. SharedMailboxesReport.csv - Contains mailbox-level information.
    2. MobileDevicesReport.csv - Contains device-level information with references to their associated mailboxes.

.PREREQUISITES
    - Exchange Server 2019
    - Administrative permissions to execute Exchange cmdlets
    - Exchange Management Shell or imported Exchange module in PowerShell
    - PowerShell 5.1 or later

.USAGE
    1. Open the Exchange Management Shell with administrative privileges.
    2. Copy and paste the script into the shell or save it as `Get-SharedMailboxesMobileAccessReport.ps1` and execute it by navigating to its directory and running:
        .\Get-SharedMailboxesMobileAccessReport.ps1
    3. The reports will be generated at the specified export paths:
        - SharedMailboxesReport.csv
        - MobileDevicesReport.csv

.OUTPUTS
    - CSV files containing detailed reports on shared mailboxes and their associated mobile devices.

.EXAMPLE
    .\Get-SharedMailboxesMobileAccessReport.ps1

.NOTES
    - Ensure that the executing account has the necessary permissions for both `Get-MobileDevice` and `Get-MobileDeviceStatistics` cmdlets.
    - Modify the `$mailboxesExportPath` and `$devicesExportPath` variables as needed to specify different export locations.

#>

# Exchange Server 2019 PowerShell Script
# Description: Retrieves shared mailboxes, checks if ActiveSync is enabled,
# identifies mobile devices accessing them, retrieves detailed statistics,
# and exports the information to separate CSV files for mailboxes and devices.

# Define the output CSV file paths
$mailboxesExportPath = "C:\Reports\SharedMailboxesReport.csv"
$devicesExportPath    = "C:\Reports\MobileDevicesReport.csv"

# Ensure the export directory exists; create it if it doesn't
$exportDirectory = Split-Path -Path $mailboxesExportPath
if (!(Test-Path -Path $exportDirectory)) {
    New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
}

# Retrieve all shared mailboxes
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Initialize arrays to store the report data
$mailboxesReport = @()
$devicesReport    = @()

# Iterate through each shared mailbox
foreach ($mailbox in $sharedMailboxes) {
    try {
        # Get Client Access settings for the mailbox
        $casMailbox = Get-CASMailbox -Identity $mailbox.Identity

        # Check if ActiveSync is enabled
        $activeSyncEnabled = $casMailbox.ActiveSyncEnabled

        # Initialize variables for mobile device information
        $mobileDeviceCount = 0

        if ($activeSyncEnabled) {
            try {
                # Retrieve mobile devices associated with the mailbox
                $mobileDevices = Get-MobileDevice -Mailbox $mailbox.Identity -ErrorAction Stop

                if ($mobileDevices) {
                    $mobileDeviceCount = $mobileDevices.Count

                    foreach ($device in $mobileDevices) {
                        try {
                            # Retrieve statistics for each mobile device
                            $deviceStats = Get-MobileDeviceStatistics -Identity $device.Identity -ErrorAction Stop

                            # Create a custom object with desired properties
                            $deviceDetail = [PSCustomObject]@{
                                MailboxName             = $mailbox.DisplayName
                                MailboxIdentity         = $mailbox.Identity
                                DeviceName              = $device.DeviceName
                                DeviceType              = $device.DeviceType
                                DeviceOS                = $device.DeviceOS
                                LastSuccessSync         = $deviceStats.LastSuccessSync
                                LastSyncAttemptTime     = $deviceStats.LastSyncAttemptTime
                                LastPolicyUpdate        = $deviceStats.LastPolicyUpdate
                                LastAccessTime          = $deviceStats.LastAccessTime
                                DeviceUserAgent         = $deviceStats.UserAgent
                                DeviceAccessState       = $deviceStats.DeviceAccessState
                                DeviceAccessStateReason = $deviceStats.DeviceAccessStateReason
                            }

                            # Add the device detail to the devices report array
                            $devicesReport += $deviceDetail
                        }
                        catch {
                            # Handle errors when retrieving device statistics
                            $deviceDetail = [PSCustomObject]@{
                                MailboxName             = $mailbox.DisplayName
                                MailboxIdentity         = $mailbox.Identity
                                DeviceName              = $device.DeviceName
                                DeviceType              = $device.DeviceType
                                DeviceOS                = $device.DeviceOS
                                LastSuccessSync         = "Error: $_"
                                LastSyncAttemptTime     = "Error: $_"
                                LastPolicyUpdate        = "Error: $_"
                                LastAccessTime          = "Error: $_"
                                DeviceUserAgent         = "Error: $_"
                                DeviceAccessState       = "Error: $_"
                                DeviceAccessStateReason = "Error: $_"
                            }

                            $devicesReport += $deviceDetail
                        }
                    }
                }
            }
            catch {
                # Handle cases where Get-MobileDevice might fail
                Write-Warning "Error retrieving mobile devices for mailbox '$($mailbox.Identity)': $_"
            }
        }

        # Create a custom object with the desired mailbox properties
        $mailboxReportObject = [PSCustomObject]@{
            MailboxName       = $mailbox.DisplayName
            MailboxIdentity   = $mailbox.Identity
            ActiveSyncEnabled = $activeSyncEnabled
            MobileDeviceCount = $mobileDeviceCount
        }

        # Add the mailbox object to the mailboxes report array
        $mailboxesReport += $mailboxReportObject
    }
    catch {
        # Handle any errors encountered while processing the mailbox
        Write-Warning "Failed to process mailbox '$($mailbox.Identity)': $_"
    }
}

# Export the mailboxes report to a CSV file
try {
    $mailboxesReport | Export-Csv -Path $mailboxesExportPath -NoTypeInformation -Encoding UTF8
    Write-Output "Mailboxes report successfully exported to '$mailboxesExportPath'."
}
catch {
    Write-Error "Failed to export mailboxes report to CSV: $_"
}

# Export the devices report to a CSV file
try {
    $devicesReport | Export-Csv -Path $devicesExportPath -NoTypeInformation -Encoding UTF8
    Write-Output "Mobile devices report successfully exported to '$devicesExportPath'."
}
catch {
    Write-Error "Failed to export mobile devices report to CSV: $_"
}
