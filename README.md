# SidexisIntegrationSetup

SidexisIntegrationSetup is a Windows console application used to configure and validate the TidyClinic–⁠Sidexis integration on a workstation.

It ensures that required components are installed, registered, and properly linked so that communication with Sidexis can function correctly.

This tool is intended to be run during initial setup or installation.

## Requirements

- Windows OS
- Administrator privileges

The application must be run as Administrator.

## Setup

Setup instructions and executable files can be downloaded [here](https://siliconavenue-my.sharepoint.com/:f:/g/personal/stephen_parinas_tidyint_com/IgDITqb1ZURNQKXvXIboD-3fAfwucXW6koxSA84sbJg6Zgs?e=kbr0ko).

The setup executable must remain in the provided folder structure alongside the required integration components.
Changing the directory layout may cause the setup to fail.

## How It Works

The setup process:
1. Verifies the program is running with elevated permissions.
2. Launches the required integration executables.
3. Confirms that the custom URI protocol has been registered.
4. Ensures Sidexis linkage is properly configured.
5. Validates that the communication mailbox exists.

If all steps complete successfully, the integration is ready for use.

## How to Run
1. Right-click the setup executable.
2. Select Run as Administrator.
3. Review console output for success or errors.

## Developer Notes

### Build Instructions

- Target framework: .NET Framework 4.8.1.
- Restore NuGet packages before building.
- Build the solution in Release mode for deployment.

### Additional Documentation

For detailed technical documentation regarding the Sidexis system integration, please refer to the official vendor documentation [here](https://www.dentsplysirona.com/en/lp/slida-partners/member-area.html).

