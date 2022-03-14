# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$displayName = $form.ChannelName
$description = $form.description
$groupId = $form.teams.GroupId

$connected = $false
try {
	$module = Import-Module MicrosoftTeams

	$pwd = ConvertTo-SecureString -string $TeamsAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $TeamsAdminUser, $pwd
	$connectTeams = Connect-MicrosoftTeams -Credential $cred
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
		if([String]::IsNullOrEmpty($description) -eq $true) {
			$channel = New-TeamChannel -displayName $displayName -groupId $groupId
		} else {
			$channel = New-TeamChannel -displayName $displayName -description $description -groupId $groupId
		}		
		Write-Information "Created Team Channel [$displayName] for Team [$groupID]"
    }
	catch
	{
		Write-error "Could not create Team Channel [$displayName]. Error: $($_.Exception.Message)"
	}
}
