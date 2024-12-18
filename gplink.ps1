param (
    [Parameter(Mandatory=$true)]
    [string]$Server,
    
    [Parameter(Mandatory=$true)]
    [string]$Username,
    
    [Parameter(Mandatory=$true)]
    [string]$Password,
    
    [Parameter(Mandatory=$true)]
    [string]$Domain,
    
    [Parameter(Mandatory=$true)]
    [string]$GPOName,
    
    [Parameter(Mandatory=$true)]
    [string]$OUDN
)

# Load the required assembly
Add-Type -AssemblyName System.DirectoryServices

# Create a new PSCredential object with the provided username and password
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)

try {
    # Connect to the Active Directory server
    $SearchBase = "CN=Policies,CN=System,$Domain"
    $SearchFilter = "(&(objectClass=groupPolicyContainer)(displayName=$GPOName))"
    $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$Server/$SearchBase", $Credential.UserName, $Credential.GetNetworkCredential().Password)
    $DirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher($DirectoryEntry)
    $DirectorySearcher.Filter = $SearchFilter
    $DirectorySearcher.PropertiesToLoad.Add("cn")
    $DirectorySearcher.PropertiesToLoad.Add("distinguishedName")
    $SearchResult = $DirectorySearcher.FindOne()

    if ($null -eq $SearchResult) {
        Write-Error "GPO '$GPOName' not found in the domain '$Domain'."
        exit 1
    }

    # Extract the GPO DN
    $GPODN = $SearchResult.Properties["distinguishedName"][0]

    Write-Host "GPO Name: $GPOName"
    Write-Host "GPO DN: $GPODN"

    # Get the current gPLink value of the OU
    $OUEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$Server/$OUDN", $Credential.UserName, $Credential.GetNetworkCredential().Password)
    $GPLink = $OUEntry.Properties["gPLink"].Value

    Write-Host "Current gPLink value: $GPLink"

    # Check if the GPO is already linked to the OU
    if ($GPLink -and $GPLink -like "*$GPODN*") {
        Write-Host "GPO '$GPOName' is already linked to the OU '$OUDN'."
        exit 0
    }

    # Append the GPO DN to the gPLink value
    $NewGPLinkEntry = "[LDAP://$GPODN;0]"
    if ($GPLink) {
        $GPLink = $GPLink + "," + $NewGPLinkEntry
    } else {
        $GPLink = $NewGPLinkEntry
    }

    Write-Host "New gPLink value: $GPLink"

    # Update the gPLink attribute of the OU
    $OUEntry.Properties["gPLink"].Value = $GPLink
    $OUEntry.CommitChanges()

    Write-Host "GPO '$GPOName' linked to the OU '$OUDN' successfully."
}
catch {
    Write-Error "An error occurred: $_"
}
