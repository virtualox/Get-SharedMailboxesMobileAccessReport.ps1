# Exchange Server 2019 Shared Mailboxes Mobile Access Report

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Output](#output)
- [Error Handling](#error-handling)
- [Customization](#customization)
- [Contributing](#contributing)

## Overview

This PowerShell script is designed for Exchange Server 2019 environments. It retrieves all shared mailboxes, checks if ActiveSync is enabled, identifies any mobile devices accessing these mailboxes, gathers detailed statistics for each device, and exports the information into two separate CSV reports:

1. **SharedMailboxesReport.csv**: Contains mailbox-level information.
2. **MobileDevicesReport.csv**: Contains device-level information with references to their associated mailboxes.

## Features

- **Retrieve Shared Mailboxes**: Identifies all shared mailboxes within the Exchange environment.
- **Check ActiveSync Status**: Determines whether ActiveSync is enabled for each shared mailbox.
- **Identify Mobile Devices**: Lists all mobile devices accessing each shared mailbox.
- **Gather Device Statistics**: Collects detailed statistics for each device, including last access time, device type, operating system, and more.
- **Export to CSV**: Generates organized CSV reports for easy analysis and record-keeping.
- **Error Handling**: Implements robust error handling to ensure smooth execution and detailed error reporting.

## Prerequisites

- **Exchange Server 2019**: Ensure that you are running Exchange Server 2019.
- **Administrative Permissions**: The executing account must have the necessary permissions to run Exchange cmdlets and access mailbox information.
- **PowerShell**: PowerShell 5.1 or later is recommended.
- **Exchange Management Shell**: The script should be run in the Exchange Management Shell or with the Exchange module imported into your PowerShell session.

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/yourusername/ExchangeSharedMailboxesMobileAccessReport.git
    ```

2. **Navigate to the Script Directory**:
    ```bash
    cd ExchangeSharedMailboxesMobileAccessReport
    ```

3. **Ensure Execution Policy Allows Script Execution**:
    - You may need to set the execution policy to allow the script to run.
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

## Usage

1. **Open Exchange Management Shell**:
    - Launch the **Exchange Management Shell** with administrative privileges.

2. **Execute the Script**:
    - If you've saved the script as `Get-SharedMailboxesMobileAccessReport.ps1`, navigate to its directory and run:
    ```powershell
    .\Get-SharedMailboxesMobileAccessReport.ps1
    ```

    - Alternatively, you can copy and paste the script directly into the Exchange Management Shell.

3. **Monitor Execution**:
    - The script will process each shared mailbox, retrieve associated mobile devices and their statistics, and export the reports to the specified paths.

4. **Review the Reports**:
    - Navigate to the export directory (default is `C:\Reports\`) and open the generated CSV files using Excel or any CSV viewer.

## Output

### 1. SharedMailboxesReport.csv

Contains the following columns:

- **MailboxName**: Display name of the shared mailbox.
- **MailboxIdentity**: Identity (email address) of the shared mailbox.
- **ActiveSyncEnabled**: Indicates if ActiveSync is enabled (`True`/`False`).
- **MobileDeviceCount**: Number of mobile devices accessing the mailbox.

**Sample Content**:

| MailboxName       | MailboxIdentity          | ActiveSyncEnabled | MobileDeviceCount |
|-------------------|--------------------------|-------------------|-------------------|
| Shared Mailbox 1  | shared1@domain.com       | True              | 2                 |
| Shared Mailbox 2  | shared2@domain.com       | False             | 0                 |
| Shared Mailbox 3  | shared3@domain.com       | True              | 1                 |

### 2. MobileDevicesReport.csv

Contains the following columns:

- **MailboxName**: Display name of the shared mailbox.
- **MailboxIdentity**: Identity of the shared mailbox.
- **DeviceName**: Name of the mobile device.
- **DeviceType**: Type/category of the device.
- **DeviceOS**: Operating system of the device.
- **LastSuccessSync**: Timestamp of the last successful synchronization.
- **LastSyncAttemptTime**: Timestamp of the last synchronization attempt.
- **LastPolicyUpdate**: Timestamp of the last policy update.
- **LastAccessTime**: Timestamp of the last access.
- **DeviceUserAgent**: User agent string of the device.
- **DeviceAccessState**: Current access state of the device.
- **DeviceAccessStateReason**: Reason for the current access state.

**Sample Content**:

| MailboxName      | MailboxIdentity     | DeviceName | DeviceType | DeviceOS   | LastSuccessSync      | LastSyncAttemptTime   | LastPolicyUpdate     | LastAccessTime        | DeviceUserAgent | DeviceAccessState | DeviceAccessStateReason |
|------------------|---------------------|------------|------------|------------|----------------------|-----------------------|----------------------|-----------------------|-----------------|-------------------|-------------------------|
| Shared Mailbox 1 | shared1@domain.com  | iPhone     | iOS        | iOS 14.4   | 2024-04-25 10:15:00  | 2024-04-25 10:15:00   | 2024-04-25 10:00:00  | 2024-04-25 10:15:00   | SomeUserAgent   | Allowed           | No issues               |
| Shared Mailbox 1 | shared1@domain.com  | Samsung    | Android    | Android 11 | 2024-04-26 08:30:00  | 2024-04-26 08:30:00   | 2024-04-26 08:00:00  | 2024-04-26 08:30:00   | AnotherUA       | Allowed           | No issues               |
| Shared Mailbox 3 | shared3@domain.com  | Android    | Android    | Android 11 | 2024-04-27 09:45:00  | 2024-04-27 09:45:00   | 2024-04-27 09:30:00  | 2024-04-27 09:45:00   | YetAnotherUA    | Allowed           | No issues               |

*Note: If an error occurs while retrieving device statistics, the respective fields will contain error messages.*

## Error Handling

- **Script-Level Errors**: The script includes `try-catch` blocks to handle and log errors during execution. Warnings are displayed for issues encountered while processing specific mailboxes or devices.
  
- **CSV Error Reporting**: If an error occurs while retrieving device statistics, the error message is recorded directly in the `MobileDevicesReport.csv` under the relevant fields.

## Customization

- **Export Paths**: Modify the `$mailboxesExportPath` and `$devicesExportPath` variables at the top of the script to change the locations where the CSV reports are saved.

    ```powershell
    $mailboxesExportPath = "C:\Reports\SharedMailboxesReport.csv"
    $devicesExportPath    = "C:\Reports\MobileDevicesReport.csv"
    ```

- **Execution Policy**: Ensure that your PowerShell execution policy allows the script to run. You can adjust it using:

    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

- **Scheduling**: For regular reporting, consider setting up a scheduled task in Windows Task Scheduler to execute the script at desired intervals.

## Contributing

Contributions are welcome! If you have suggestions for improvements, bug fixes, or new features, feel free to open an issue or submit a pull request.

### Steps to Contribute:

1. **Fork the Repository**: Click on the "Fork" button at the top right of this page.
2. **Clone Your Fork**:
    ```bash
    git clone https://github.com/yourusername/ExchangeSharedMailboxesMobileAccessReport.git
    ```
3. **Create a New Branch**:
    ```bash
    git checkout -b feature/YourFeatureName
    ```
4. **Make Your Changes**: Edit the script or documentation as needed.
5. **Commit Your Changes**:
    ```bash
    git commit -m "Add your descriptive commit message here"
    ```
6. **Push to Your Fork**:
    ```bash
    git push origin feature/YourFeatureName
    ```
7. **Open a Pull Request**: Navigate to the original repository and click on "Compare & pull request".

---

**Disclaimer**: Use this script at your own risk. Ensure you have proper backups and have tested the script in a non-production environment before deploying it in a live setting.
