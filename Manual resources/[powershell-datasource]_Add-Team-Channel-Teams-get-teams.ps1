#Input: TeamsAdminUser
#Input: TeamsAdminPWD

$connected = $false
try {
	$module = Import-Module MicrosoftTeams
	$pwd = ConvertTo-SecureString -string $TeamsAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $TeamsAdminUser, $pwd
	$teamsConnection = Connect-MicrosoftTeams -Credential $cred
    Write-Information "Connected to Microsoft Teams"
    $connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft Teams. Error: $($_.Exception.Message)"
}

if ($connected)
{
	try {
	    $teams = Get-Team
        Write-Information "Result count: $(@($teams).Count)"

        if(@($teams).Count -gt 0){
            foreach($team in $teams)
            {
                $addRow = @{DisplayName=$team.DisplayName; Description=$team.Description; MailNickName=$team.MailNickName; Visibility=$team.Visibility; Archived=$team.Archived; GroupId=$team.GroupId;}
                Write-Output $addRow
            }
        }
	}
	catch
	{
		Write-Error "Error getting Teams. Error: $($_.Exception.Message)"
	}
}

