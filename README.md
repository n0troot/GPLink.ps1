# GPLink.ps1

This PowerShell script allows you to link a Group Policy Object (GPO) to an Organizational Unit (OU) in Active Directory using the `System.DirectoryServices` namespace. It provides a convenient way to automate the process of linking GPOs to OUs without relying on the "GroupPolicy" module.

## Prerequisites

- Active Directory environment with the necessary permissions to link GPOs to OUs.
- PowerShell execution policy set to allow running scripts.

## Usage

1. Save the script with a `.ps1` extension (e.g., `gplink.ps1`).

2. Open a PowerShell console and navigate to the directory where the script is saved.

3. Run the script with the required parameters:

   ```powershell
   .\gplink.ps1 -Server "server_ip" -Username "domain\username" -Password "password" -Domain "domain_dn" -GPOName "gpo_name" -OUDN "ou_dn"
   ```

   Replace the following placeholders with the appropriate values:
   - `server_ip`: The IP address or hostname of the Active Directory server.
   - `domain\username`: The username with sufficient permissions to link GPOs to OUs.
   - `password`: The password for the specified user account.
   - `domain_dn`: The distinguished name (DN) of your Active Directory domain.
   - `gpo_name`: The name of the GPO you want to link.
   - `ou_dn`: The distinguished name (DN) of the target OU.

   Example:
   ```powershell
   .\gplink.ps1 -Server "10.0.0.1" -Username "contoso\admin" -Password "P@ssw0rd" -Domain "DC=contoso,DC=com" -GPOName "Default Domain Policy" -OUDN "OU=Workstations,DC=contoso,DC=com"
   ```

4. The script will perform the following actions:
   - Connect to the Active Directory server using the provided credentials.
   - Search for the specified GPO by its name.
   - Retrieve the current `gPLink` value of the target OU.
   - Check if the GPO is already linked to the OU.
   - If the GPO is not linked, append the GPO DN to the `gPLink` value.
   - Update the `gPLink` attribute of the OU with the new value.

5. The script will display the GPO name, GPO DN, current `gPLink` value, new `gPLink` value, and a success message if the GPO is successfully linked to the OU.

## Error Handling

If an error occurs during the execution of the script, an error message will be displayed using the `Write-Error` cmdlet. The error message will include the specific exception details.

If the specified GPO is not found in the domain, an error message will be displayed, and the script will exit with a status code of 1.

If the GPO is already linked to the OU, a message will be displayed indicating that the GPO is already linked, and the script will exit with a status code of 0.

## Permissions

To successfully link a GPO to an OU using this script, the user account specified in the `-Username` parameter must have sufficient permissions. The required permissions may include:

- "GenericAll" permission on the GPO object.
- "Write gPLink" permission on the target OU.

If the script encounters an "Access is denied" error, it indicates that the user account lacks the necessary permissions. In such cases, it is recommended to involve your Active Directory administrator or a domain expert to investigate and resolve the permission-related issues.

## Notes

- This script uses the `System.DirectoryServices` namespace to interact with Active Directory and modify the `gPLink` attribute of the OU.
- The script assumes that you have a valid Active Directory environment and the necessary permissions to link GPOs to OUs.
- Make sure to replace the placeholder values in the script with the appropriate information specific to your environment.
- It is always recommended to test the script in a non-production environment before using it in a production setting.

## License

This script is provided as-is without any warranty. Feel free to modify and use it according to your needs.
