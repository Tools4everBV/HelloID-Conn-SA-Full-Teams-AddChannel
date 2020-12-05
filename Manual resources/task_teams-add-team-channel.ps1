$connected = $false
try {
	Import-Module MicrosoftTeams
	$pwd = ConvertTo-SecureString -string $TeamsAdminPWD -AsPlainText –Force
	$cred = New-Object System.Management.Automation.PSCredential $TeamsAdminUser, $pwd
	Connect-MicrosoftTeams -Credential $cred
    HID-Write-Status -Message "Connected to Microsoft Teams" -Event Information
    HID-Write-Summary -Message "Connected to Microsoft Teams" -Event Information
	$connected = $true
}
catch
{	
    HID-Write-Status -Message "Could not connect to Microsoft Teams. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Microsoft Teams" -Event Failed
}

if ($connected)
{
	try {
		if([String]::IsNullOrEmpty($description) -eq $true) {
			New-TeamChannel -displayName $displayName -groupId $groupId
		} else {
			New-TeamChannel -displayName $displayName -description $description -groupId $groupId
		}		
		HID-Write-Status -Message "Created Team Channel [$displayName] for Team [$groupID]" -Event Success
		HID-Write-Summary -Message "Successfully created Team Channel [$displayName] for Team [$groupID]" -Event Success
	}
	catch
	{
		HID-Write-Status -Message "Could not create Team Channel [$displayName]. Error: $($_.Exception.Message)" -Event Error
		HID-Write-Summary -Message "Failed to create Team Channel [$displayName]" -Event Failed
	}
}