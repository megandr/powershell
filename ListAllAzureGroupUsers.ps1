# Requires the Microsoft.Graph PowerShell module
# Install if needed: Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

# Get all groups
$groups = Get-MgGroup -All

# Create an array to store results
$results = @()

foreach ($group in $groups) {
    Write-Host "Processing group: $($group.DisplayName)" -ForegroundColor Cyan
    
    try {
        # Get members of each group
        $members = Get-MgGroupMember -GroupId $group.Id -All
        
        foreach ($member in $members) {
            # Get additional details for each member
            $memberDetails = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
            if ($memberDetails) {
                $results += [PSCustomObject]@{
                    GroupDisplayName = $group.DisplayName
                    GroupId = $group.Id
                    GroupDescription = $group.Description
                    MemberDisplayName = $memberDetails.DisplayName
                    MemberEmail = $memberDetails.UserPrincipalName
                    MemberType = "User"
                }
            } else {
                # If not a user, might be a service principal or group
                $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $member.Id -ErrorAction SilentlyContinue
                if ($servicePrincipal) {
                    $results += [PSCustomObject]@{
                        GroupDisplayName = $group.DisplayName
                        GroupId = $group.Id
                        GroupDescription = $group.Description
                        MemberDisplayName = $servicePrincipal.DisplayName
                        MemberEmail = $servicePrincipal.AppId
                        MemberType = "ServicePrincipal"
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Error processing group $($group.DisplayName): $_"
    }
}

# Export results to CSV
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$csvPath = "AzureADGroupMembers_$timestamp.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "`nExported results to: $csvPath" -ForegroundColor Green
Write-Host "Total groups processed: $($groups.Count)" -ForegroundColor Green
Write-Host "Total members found: $($results.Count)" -ForegroundColor Green

# Disconnect from Microsoft Graph
Disconnect-MgGraph