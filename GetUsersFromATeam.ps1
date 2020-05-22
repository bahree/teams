#Global variables to set:
#path of the file where to export
#specific ID of the team that you want the users for. 
$exportLocation = "C:\temp\RL-decision-AI-export.csv"
$TEAM_ID = "f3f9ad1f-beea-4026-9b86-dd3788404999"
            
$Report = @()

# counters
$ownerCount = 0
$memberCount = 0
$guestCount = 0

#connect to teams
Connect-MicrosoftTeams

$team = Get-Team -GroupId $TEAM_ID

#Patience, supposed to be a virtue
Write-Host -ForegroundColor Blue "Successfully connected to Team: $($team.DisplayName)"
Write-Host -ForegroundColor Blue "Getting all users in the team"
Write-Host -ForegroundColor Blue "Please be patient, if there are a lot of users, this can take a while..."

# Attempt to get team users, throw error message if no access
try {
	# Get team members
	$users = Get-TeamUser -GroupId $team.groupID

	# Loop through and get all the users
	$currentIndex = 1
	
	# foreach user create a line in the report
	ForEach($user in $users) {
		# Show some output to the user
		Write-Progress -Id 0 -Activity "Generating user report from Teams" -Status "$currentIndex of $($users.Count)" -PercentComplete (($currentIndex / $users.Count) * 100)
	
		# Maintain a count of user types
		switch($user.Role) {
			"owner" { $ownerCount++ }
			"member" { $memberCount++ }
			"guest" { $guestCount++ }
		}

		# Create an object to hold all values
		$ReportObject = New-Object PSObject -Property @{
			User = $user.Name
			Email = $user.User
			Role = $user.Role
		}

		# Add to the report
		$Report += $ReportObject
		
		$currentIndex++
	}
} 
catch [Microsoft.TeamsCmdlets.PowerShell.Custom.ErrorHandling.ApiException] {
	Write-Host -ForegroundColor Yellow "No access to $($team.DisplayName) team, cannot generate report"
}

#Complete progress
Write-Progress -Id 0 -Activity " " -Status " " -Completed

# Disconnect from the teams
Disconnect-MicrosoftTeams

# Write out details for the user
Write-Host -ForegroundColor Green "============================================================"
Write-Host -ForegroundColor Green "                Microsoft Teams User Report                 "
Write-Host -ForegroundColor Green ""
Write-Host -ForegroundColor Green "Team Details:"
Write-Host -ForegroundColor Green "Name: $($team.DisplayName)"
Write-Host -ForegroundColor Green "Description: $($team.Description)"
Write-Host -ForegroundColor Green "Mail Nickname: $($team.MailNickName)"
Write-Host -ForegroundColor Green "Archived: $($team.Archived)"
Write-Host -ForegroundColor Green "Visiblity: $($team.Visibility)"
Write-Host -ForegroundColor Green ""
Write-Host -ForegroundColor Green "Team User Details:"
Write-Host -ForegroundColor Green "Owners - $($ownerCount)"
Write-Host -ForegroundColor Green "Members - $($memberCount)"
Write-Host -ForegroundColor Green "Guests - $($guestCount)"
Write-Host -ForegroundColor Green "============================================================"

$Report | Export-CSV $exportLocation -NoTypeInformation -Force
Write-Host -ForegroundColor Blue "Exported report to $($exportLocation)"
