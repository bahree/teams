#set these variables, to what makes sense in your situation. The email here is the one that is the one connected to your teams account.
$exportLocation = "C:\temp\team-details.csv"
$emailAddress = "your-email@shouldbeputhere.com"

Connect-MicrosoftTeams

#Patience
Write-Host -ForegroundColor Blue "Successfully connected to Teams"
Write-Host -ForegroundColor Blue "Getting all team details for user: $($emailAddress)"
Write-Host -ForegroundColor Blue "Please be patient, if there are a lot of teams, this can take a while..."

# Get all of the team Groups IDs
# $GetUsersTeams = (Get-Team).GroupID
$GetUsersTeams = Get-Team -User $emailAddress

$Report = @()

# Will hold a basic count of user types and teams
$unavailableTeamCount = 0

# Loop through all teams that the user belongs to
$currentIndex = 1

ForEach($thisTeam in $GetUsersTeams) {
	# Show some output to the user
    Write-Progress -Id 0 -Activity "Building report from Microsoft Teams" -Status "$currentIndex of $($GetUsersTeams.Count)" -PercentComplete (($currentIndex / $GetUsersTeams.Count) * 100)
	
    # Attempt to get team details, throw error message if no access
    try {
        # Get team members
        #$users = Get-TeamUser -GroupId $thisTeam.groupID

		# Create an object to hold all values
        $teamReportObject = New-Object PSObject -Property @{
                GroupID = $thisTeam.GroupID
				TeamName = $thisTeam.DisplayName
                Description = $thisTeam.Description
                Archived = $thisTeam.Archived
                Visibility = $thisTeam.Visibility
				eMail = $thisTeam.MailNickName
            }

            # Add to the report
            $Report += $teamReportObject
       
    } catch [Microsoft.TeamsCmdlets.PowerShell.Custom.ErrorHandling.ApiException] {
        Write-Host -ForegroundColor Yellow "No access to $($team.DisplayName) team, cannot generate report"
        $unavailableTeamCount++
    }
	
	$currentIndex++
}
Write-Progress -Id 0 -Activity " " -Status " " -Completed

# Disconnect from the teams service
Disconnect-MicrosoftTeams

# Provide some nice output
Write-Host -ForegroundColor Green "============================================================"
Write-Host -ForegroundColor Green "                Microsoft Teams User Report                 "
Write-Host -ForegroundColor Green ""
Write-Host -ForegroundColor Green "  Count of All Teams - $($GetUsersTeams.Count)                "
Write-Host -ForegroundColor Green "  Count of Inaccesible Teams - $($unavailableTeamCount)         "
Write-Host -ForegroundColor Green ""

$Report | Export-CSV $exportLocation -NoTypeInformation -Force
Write-Host -ForegroundColor Blue "Exported report to $($exportLocation)"
