# LSP Onboarding Automator

A GUI-driven PowerShell automation tool designed to streamline Microsoft 365 user provisioning for Legacy Service Partners (LSP). This script eliminates manual portal navigation by automatically handling Entra ID account creation, manager assignment, explicit group-based licensing, and complex Exchange Online alias routing.

## Features
* **Standardized GUI:** Provides a simple Windows Forms interface for inputting new hire data.
* **Automated Entra ID Provisioning:** Generates passwords, formats UPNs, and checks for duplicate accounts (automatically appending `01` to resolve conflicts).
* **Manager Linking:** Automatically resolves and links the user's manager via Microsoft Graph.
* **Exchange Routing:** Waits for mailbox sync and applies the primary SMTP and all required alias domains based on the partner's specific routing rules.
* **ServiceNow Integration:** Outputs a pristine, formatted block of work notes ready to be copied directly into the onboarding ticket resolution.

## Prerequisites
This tool is designed to run in **Windows PowerShell 5.1** to ensure compatibility between the Windows Forms GUI and Microsoft authentication brokers. 

The following PowerShell modules are required:
* `Microsoft.Graph` (v2.0 or newer)
* `ExchangeOnlineManagement`

To install the required modules on a new machine, run the following command in an administrative Windows PowerShell terminal:

    Install-Module Microsoft.Graph, ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber

## Repository Structure

    LSP-Onboarding/
    ├── LSP-Onboarding.ps1       # The primary execution script and GUI
    ├── README.md                # Documentation
    └── Config/
        └── Companies.json       # Partner routing rules and logic

## Usage
1. Clone or download the repository to your local machine or secure workspace.
2. Right-click `LSP-Onboarding.ps1` and select **Run with PowerShell**.
3. Authenticate with your administrative credentials when prompted by the Microsoft login windows.
4. Fill out the new hire details in the GUI.
5. Click **Provision New User**.
6. Wait for the success output in the log box, then copy the resolution block into your work ticket.

## Configuration (Companies.json)
All partner-specific logic is centralized in `Config\Companies.json`. If a new LSP partner is acquired, simply add a new JSON object to the file. 

**Required Keys:**
* `CompanyName`: The official company string to be injected into the Entra ID profile.
* `EmailFormat`: The naming convention used for the local part of the email address (e.g., `FirstInitialLastName`, `FirstName.LastName`).
* `PrimaryDomain`: The default tenant routing domain (e.g., `hersandhis.com`).
* `AliasDomains`: An array of strings representing any additional domains that need to be attached to the mailbox (e.g., `["hersandhisplumbing.com"]`). Leave empty `[]` if none apply.
