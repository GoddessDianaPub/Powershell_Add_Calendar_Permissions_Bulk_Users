# Define Source UPN and Access Rights
$sourceUPN = "user@domain.com"
$accessRights = @("accessRight")

# Define the users as an array
$DestinationUsers = @(
    "UPN",
    "User"
)

# Loop through each secretary and configure permissions
foreach ($DestinationUser in $DestinationUsers) {
    $calendarIdentity = "${sourceUPN}:\Calendar"

    # Add permissions
    Write-Host "Granting $accessRights access to $DestinationUsers for $calendarIdentity..."
    try {
        Add-MailboxFolderPermission -Identity $calendarIdentity -User $DestinationUser -AccessRights $accessRights
        Write-Host "Successfully granted $accessRights to $DestinationUsers." -ForegroundColor Green
    } catch {
        Write-Host "Failed to grant permissions to $DestinationUsers. Error: $_" -ForegroundColor Red
    }
}

Get-MailboxFolderPermission -Identity $calendarIdentity | Format-Table FolderName, User, AccessRights -AutoSize