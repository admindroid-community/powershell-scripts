<#
=============================================================================================
Name:        Manage Microsoft 365 Groups Using All-in-One PowerShell Script
Description: Performs 19 Microsoft 365 group management actions, including single and bulk operations, to save time and reduce effort. 
Version:     1.0
Website:     o365reports.com

~~~~~~~~~~~~~~~~~~
Script Highlights:
~~~~~~~~~~~~~~~~~~
1. Allows you to perform 19 actions to manage M365 groups.  
2. Supports 7 single management actions through prompts.
3. Performs 12 bulk operations that require a CSV input file.
4. Enables you to perform multiple actions without running the script repeatedly. 
5. Automatically installs the Microsoft Graph PowerShell module if it is not already available. 
6. Automatically exports an output CSV log after each execution for easy tracking and analysis. 
7. Supports Certificate-based authentication (CBA) as well. 
8. The script can be executed with an MFA enabled account too. 
9. The script is scheduler-friendly, making it easy to automate tasks. 
 
For detailed script execution: https://o365reports.com/microsoft-365-group-management-using-powershell/  

=============================================================================================
#>

Param(
	[int]$Action,
	[switch]$MultipleActionsMode,
    [string]$InputCsvFilePath,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)
Function Connect_MgGraph {
    #Check for module installatiion
    $Module = Get-Module -Name microsoft.graph -ListAvailable
    if($Module.Count -eq 0){
        Write-Host "`nMicrosoft Graph PowerShell SDK is not available"  -ForegroundColor yellow 
        $Confirm = Read-Host Are you sure want to install the module? [Y]Yes [N]No
        if($Confirm -match '[Yy]'){
            Write-Host "Installing Microsoft Graph PowerShell module..."
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else{
            Write-Host "Microsoft Graph PowerShell module is required to run this script. Please install module using Install-Module Microsoft.Graph cmdlet." 
            Exit
        }
    }
    #Disconnect Existing MgGraph session
    if ($CreateSession.IsPresent) {
    	Disconnect-MgGraph | Out-Null
    }
    Write-Host "`nConnecting to Microsoft Graph..." 
    if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
    {  
        Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    else{
        Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All" ,"Organization.Read.All","RoleManagement.ReadWrite.Directory" -NoWelcome
    }
}

Function Show-Menu {
	Write-Host "`n=======================================================" -ForegroundColor Yellow
    Write-Host "`tManage M365 Group Management" -ForegroundColor Green
	Write-Host "=======================================================`n" -ForegroundColor Yellow
	Write-Host "`t~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Yellow
	Write-Host "`t  Single Operations" -ForegroundColor Cyan
	Write-Host "`t~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Yellow
	Write-Host @"
	 1. Create group
	 2. Add teams to the group
	 3. Assign licenses to the group
	 4. Remove licenses from the group
	 5. Remove user from all group roles
	 6. Delete group
	 7. Restore group 
	 
"@ -ForegroundColor  Yellow
	 Write-Host "`t~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Yellow
	 Write-Host "`t  Bulk Operations (input CSV)" -ForegroundColor Cyan
	 Write-Host "`t~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" -ForegroundColor Yellow
	 Write-Host @"
	 8. Create bulk groups
	 9. Add user to bulk groups	 
	10. Add bulk users to a group
	11. Add bulk users to bulk groups
	12. Add teams to bulk groups
	13. Assign the licenses to the groups	 
	14. Remove the licenses from the groups	 
	15. Remove users from all group roles
	16. Remove user from specific groups
	17. Remove users from specific groups	 
	18. Delete bulk groups	 
	19. Restore bulk groups	
	 0. Exit

"@ -ForegroundColor Yellow
	Write-Host "=======================================================`n" -ForegroundColor Yellow
}
Function Create_Group {
	param(
		[string]$DisplayName,
		[string]$MailNickName,
		[string[]]$GroupOwnerUPN,
		[string[]]$MembersUPN,
		[string]$Description = "",
		[string]$Privacy,
		[bool]$isRoleAssignable = $false
	)
	try {
		$groupBody = @{
			DisplayName = $DisplayName
			MailEnabled = $true
			MailNickName= $MailNickName
			SecurityEnabled=$true
			GroupTypes  = @("Unified")
			"Owners@odata.bind" = @(
				$GroupOwnerUPN | ForEach-Object {
					"https://graph.microsoft.com/v1.0/users/$($_.Trim())"
				}
			)
		}
		if ($MembersUPN.Count -gt 0) {
			$groupBody["members@odata.bind"] = @(
				$MembersUPN | ForEach-Object {
					"https://graph.microsoft.com/v1.0/users/$($_.Trim())"
				}
			)
		}
		if($Description -ne ""){
			$groupBody["Description"] = $Description
		}
		else{
			$Description = " -"
		}
		if(($Privacy -eq "Public") -or ($Privacy -eq "")){
			$groupBody["Visibility"] = "Public"
		}
		else{
			$groupBody["Visibility"] = "Private"
		}

		if($isRoleAssignable -and ($Privacy -eq "Private")){
			$groupBody["isAssignableToRole"] = $true
		}
		else{
			$groupBody["isAssignableToRole"] = $false
		}
		New-MgGroup -BodyParameter $groupBody -ErrorAction stop | Out-Null
		Write-Log -EventTime (Get-Date) -GroupName $DisplayName -Action "Create Group" -Value " -" -Status "Success" -Err " -"
		if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			Write-Host "`nGroup $($DisplayName) created successfully." -ForegroundColor Green
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $DisplayName -Action "Create Group" -Value " -" -Status "Failed" -Err $_.Exception
		if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			Write-Host "`nGroup $($DisplayName) creation failed. Error: $($_.Exception)" -ForegroundColor Red
		}
	}
}

Function Create_Single_Group{
	$DisplayName = (Read-Host "`nEnter the displayName of the group").Trim()
	$Description = (Read-Host "Enter the group description(optional)").Trim()
	$MailNickName = (Read-Host "Enter the MailNickName of the group").Trim()
	$GroupOwnerUPN = (Read-Host "Enter the group owner UPNs (comma separated)") -split "," | ForEach-Object Trim | Where-Object Length
	$AddMember = (Read-Host "Do you want to assign members to this group? [Y] Yes [N] No").Trim()
	if($AddMember -match '[Yy]'){
		$MembersUPN = (Read-Host "Enter member UPNs (comma separated)") -split "," | ForEach-Object Trim | Where-Object Length
	}
	$Privacy = (Read-Host "Enter the group privacy").Trim()
	if($Privacy -eq "Private"){
		$RoleAssignment = (Read-Host "Do you need role assignment for the group? [Y]Yes [N]No").Trim()
		if($RoleAssignment -match '[Yy]'){
			$isRoleAssignable = $true
		}
		else{
			$isRoleAssignable = $false
		}
	}
	else{
		$Privacy = "Public"
		$isRoleAssignable = $false
	}
	Create_Group -DisplayName $DisplayName -Description $Description -GroupOwnerUPN $GroupOwnerUPN -MailNickName $MailNickName -MembersUPN $MembersUPN -Privacy $Privacy -isRoleAssignable $isRoleAssignable
}
Function Create_Bulk_Groups {
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$DisplayName = $_.DisplayName.Trim()
			$MailNickName = $_.MailNickName.Trim()
			if($_.Description -and $_.Description.Trim()){
				$Description = $_.Description.Trim()
			}
			else{
				$Description = ""
			}
			$GroupOwnerUPN = $_.GroupOwnerUPN -split "," | ForEach-Object Trim | Where-Object Length
			if($null -ne $_.MembersUPN){
				$MembersUPN = $_.MembersUPN -split "," | ForEach-Object Trim | Where-Object Length
			}
			$Privacy = $_.Privacy.Trim()
			if($_.Privacy -eq "Private"){
				if($_.isRoleAssignable -and $_.isRoleAssignable -match '[Yy]'){
					$isRoleAssignable = $true
				}
				else{
					$isRoleAssignable = $false
				}
			}
			else{
				$isRoleAssignable = $false
			}
			Write-Progress  -Activity "Processed group count: $($Count)" ` -Status "Currently processing: $($DisplayName)" -PercentComplete 100
			$Count++
			Create_Group -DisplayName $DisplayName -Description $Description -GroupOwnerUPN $GroupOwnerUPN -MailNickName $MailNickName -MembersUPN $MembersUPN -Privacy $Privacy -isRoleAssignable $isRoleAssignable
		}	
	}
	catch {
		Write-Log  -EventTime (Get-Date) -Action "Create group" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally {
		Write-Progress -Activity "Completed" -Completed
	}
}

Function Delete_Group{
	param(
		[string]$MailAddress,
		[string]$DeleteType
	)
	try {
		$Group = Get-MgGroup -Filter "mail eq '$MailAddress'"
		Remove-MgGroup -GroupId $Group.Id -Confirm:$false -ErrorAction Stop
		if($DeleteType -eq "Hard"){
			Start-Sleep -Seconds 5
			Remove-MgDirectoryDeletedItem -DirectoryObjectId $Group.Id -ErrorAction stop
			Write-Log  -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Permanently delete the group" -Value " -" -Status "Success" -Err " -"
			if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
				Write-Host "`nGroup $($Group.DisplayName) deleted permanently." -ForegroundColor Green
			}
		}
		else{
			Write-Log  -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Move group to deleted items" -Value " -" -Status "Success" -Err " -"
			if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
				Write-Host "`nGroup $($Group.DisplayName) was successfully moved to the deleted items list." -ForegroundColor Green
			}
		}
	}
	catch {
		if($DeleteType -eq "Hard"){
			Write-Log -EventTime(Get-Date) -GroupName $Group.DisplayName -Action "Permanently delete the group" -Value " -" -Status "Failed" -Err $_.Exception
			if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
				Write-Host "`nFailed to permanently remove the group $($Group.DisplayName). Error: $($_.Exception)" -ForegroundColor Red
			}
		}
		else{
			Write-Log -EventTime(Get-Date) -GroupName $Group.DisplayName -Action "Move group to deleted items" -Value " -" -Status "Failed" -Err $_.Exception
			if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
				Write-Host "`nFailed to move the group $($Group.DisplayName) to the deleted items list. Error: $($_.Exception)" -ForegroundColor Red
			}
		}
	}
}

Function Delete_Single_Group{
	$MailAddress =(Read-Host "`nEnter the mail address of the group").Trim()
	$DeleteType = (Read-Host "Enter deletion type of the action (Soft/Hard)").Trim()
	Delete_Group -MailAddress $MailAddress -DeleteType $DeleteType
}

Function Delete_Bulk_Groups{
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$MailAddress = $_.MailAddress.Trim()
			if ($_.DeleteType -and $_.DeleteType.Trim()) {
				$DeleteType = $_.DeleteType.Trim()
			} else {
				$DeleteType = "Soft"
			}
			Write-Progress -Activity "Processed Group count:$($Count)" -Status "Currently processing: $($MailAddress)" -PercentComplete 100
			$Count++
			Delete_Group -MailAddress $MailAddress -DeleteType $DeleteType
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Delete a group" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally {
		Write-Progress -Activity "Completed" -Completed
	}
}

Function Restore_Group{
	param(
		[string]$MailAddress
	)
	try {
		[bool]$isNotRestored = $true
		(Get-MgDirectoryDeletedItemAsGroup  -all) | ForEach-Object{
			if($_.Mail -eq $MailAddress){
				Restore-MgDirectoryDeletedItem -DirectoryObjectId $_.Id -ErrorAction Stop | Out-Null
				$isNotRestored = $false
				Write-Log -EventTime (Get-Date) -GroupName $_.DisplayName -Action "Restore group" -Value " -" -Status "Success" -Err " -"
				if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
					Write-Host "`nThe group $($_.DisplayName) was successfully restored." -ForegroundColor Green
				}
			}
		}
		if($isNotRestored){
			Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Restore group" -Value " -" -Status "Failed" -Err "Group not found in the deleted items."
			if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
				Write-Host "`nFailed to restore the group $($MailAddress). Error: The group $($MailAddress) was not found in the deleted items list." -ForegroundColor Red
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Restore group" -Value " -" -Status "Failed" -Err $_.Exception
		if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			Write-Host "`nFailed to restore the group $($MailAddress). Error: $($_.Exception)" -ForegroundColor Red
		}
	}
}

Function Restore_Single_Group{
	$MailAddress = (Read-Host "`nEnter the group's mail address to restore it").Trim()
	Restore_Group -MailAddress $MailAddress
}

Function Restore_Bulk_Groups {
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease Enter the input file path "
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object { 
			$MailAddress = $_.MailAddress.Trim()
			Write-Progress -Activity "Processed Group count:$($Count)" -Status "Currently Processing: $($MailAddress)" -PercentComplete 100
			$Count++
			Restore_Group -MailAddress $MailAddress
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Restore the groups" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

Function Add_Team_To_Group{
	param(
		[string]$MailAddress
	)
	try {
		$Group = Get-MgGroup -Filter "mail eq '$MailAddress'"
		$body = @{
			"template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
			"group@odata.bind"    = "https://graph.microsoft.com/v1.0/groups/$($Group.Id)"
		}
		New-MgTeam -BodyParameter $body -ErrorAction Stop
		Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add Teams to group" -Value " -" -Status "Success" -Err " -"
		if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			Write-Host "`nSuccessfully added Teams to the group $($Group.DisplayName)." -ForegroundColor Green
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add Teams to group" -Value " -" -Status "Failed" -Err $_.Exception
		if([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			Write-Host "`nFailed to add Teams to the group $($Group.DisplayName). Error: $($_.Exception)" -ForegroundColor Red
		}
	}
}

Function Add_Team_To_Single_Group{
	$MailAddress = (Read-Host "`nEnter the mail address of the group (Group owners must have a Teams license)").Trim()
	Add_Team_To_Group -MailAddress $MailAddress
}

Function Add_Team_To_Bulk_Group{
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path (Group owners must have a Teams license)"
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object { 
			$MailAddress = $_.MailAddress.Trim()
			Write-Progress -Activity "Processed group count: $($Count)" -Status "Currently Processing $($MailAddress)" -PercentComplete 100
			$Count++
			Add_Team_To_Group -MailAddress $MailAddress
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Add Teams to the groups" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

Function Add_UserToBulkGroups {
	$UPN = (Read-Host "`nPlease Enter the User Principal Name").Trim()
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "Please Enter the input file path "
	}
	try {
		$Count = 1
		$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop ).Id
		Import-Csv -Path $InputCsvFilePath | ForEach-Object { 
			$MailAddress = $_.MailAddress.Trim()
			$Membership = $_.Membership.Trim()
			Write-Progress  -Activity "Processed groups count: $Count" ` -Status "Currently Processing: $($MailAddress)" -PercentComplete 100
			$Count++
			try {
				$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
				if($Membership -eq "owner"){
					New-MgGroupOwnerByRef -GroupId $Group.Id -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserId"} -ErrorAction Stop
				}
				elseif($Membership -eq "member"){
					New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		} 
    }
    catch { 
		Write-Log -EventTime (Get-Date) -Action "Add user to groups" -Value " -" -Status "Failed" -Err $_.Exception
    }
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Add_BulkUsersToSingleGroup {
	$MailAddress = (Read-Host "`nPlease enter the Mail Address of the Group").Trim()
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "Please Enter the input file path "
	}
	try {
		$Count = 1
		$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{ 
			$UPN = $_.UPN.Trim()
			$Membership = $_.Membership.Trim()
			try {
				Write-Progress  -Activity "Processed users count: $Count" ` -Status "Currently Processing: $($UPN)" -PercentComplete 100
				$Count++
				$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
				if($Membership -eq "owner"){
					New-MgGroupOwnerByRef -GroupId $Group.Id -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserId"} -ErrorAction Stop
				}
				elseif($Membership -eq "member"){
					New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Add users to group" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Add_BulkUsersToBulkGroups{
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$Count = 1
		Import-Csv $InputCsvFilePath | ForEach-Object {
			$UPN = $_.UPN.Trim()
			$MailAddress = $_.MailAddress.Trim()
			$Membership = $_.Membership.Trim()
			Write-Progress  -Activity "Processed group count: $Count" ` -Status "Currently Processing: $($MailAddress)" -PercentComplete 100
			$Count++
			try {
				$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
				$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
				if($Membership -eq "owner"){
					New-MgGroupOwnerByRef -GroupId $Group.Id -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserId"} -ErrorAction Stop
				}
				if($MemberShip -eq "member"){
					New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Add $($Membership) to group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
	}
	catch {
		Write-Log -EventTime(Get-Date) -Action "Add bulk users to groups" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Remove_UserFromAllGroups{
	$UPN = (Read-Host "`nPlease enter the UPN").Trim()
	try {
		$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
		$MemberGroups = Get-MgUserMemberOfAsGroup -UserId $UserId | Where-Object{
			$_.GroupTypes -contains "Unified" -and $_.GroupTypes -notcontains "DynamicMembership"
		} -ErrorAction Stop
		$OwnedGroups = Get-MgUserOwnedObjectAsGroup -UserId $UserId | Where-Object{
			$_.GroupTypes -contains "Unified" -and $_.GroupTypes -notcontains "DynamicMembership"
		} -ErrorAction Stop
		$Count = 1
		foreach($Group in $MemberGroups){
			try {
				Write-Progress -Activity "Processed group count:$Count" -Status "Currently Processing: $($Group.DisplayName)" -PercentComplete 100
				$Count++
				Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove member from group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove member from group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
		foreach($Group in $OwnedGroups){
			try {
				Write-Progress -Activity "Processed group count:$Count" -Status "Currently Processing: $($Group.DisplayName)" -PercentComplete 100
				$Count++
				Remove-MgGroupOwnerByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove owner from group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove owner from group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
		Write-Host "`nSuccessfully removed the user $($UPN) from all group roles. Details are available in the log file" -ForegroundColor Green
	}
	catch {
		Write-Log -EventTime (Get-Date)  -Action "Remove user from all group roles" -Value " -" -Status "Failed" -Err $_.Exception
		Write-Host "`nFailed to remove the user $($UPN) from all group roless. Error : $($_.Exception)" -ForegroundColor Red
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Remove_UsersFromAllGroups {
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$UPN = $_.UPN.Trim()
			Write-Progress  -Activity "Processed users count: $($Count)" ` -Status "Currently Processing: $($UPN)" -PercentComplete 100
			$Count++
			try {
				$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
				$MemberGroups = Get-MgUserMemberOfAsGroup -UserId $UserId | Where-Object{
					$_.GroupTypes -contains "Unified" -and $_.GroupTypes -notcontains "DynamicMembership"
				} -ErrorAction Stop
				foreach($Group in $MemberGroups){
					try {
						Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
						Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove member from group" -Value $UPN -Status "Success" -Err " -"
					}
					catch {
						Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove member from group" -Value $UPN -Status "Failed" -Err $_.Exception
					}
				}
				$OwnedGroups = Get-MgUserOwnedObjectAsGroup -UserId $UserId | Where-Object{
					$_.GroupTypes -contains "Unified" -and $_.GroupTypes -notcontains "DynamicMembership"
				} -ErrorAction Stop
				foreach($Group in $OwnedGroups){
					try {
						Remove-MgGroupOwnerByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
						Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove owner from group" -Value $UPN -Status "Success" -Err " -"
					}
					catch {
						Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove owner from group" -Value $UPN -Status "Failed" -Err $_.Exception
					}
				}
			}
			catch {
				Write-Log -EventTime (Get-Date) -Action "Remove users from all group roles" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Remove users from all group roles" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Remove_UserFromListOfGroups {
	$UPN = (Read-Host "`nPlease enter the UPN").Trim()
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "Please enter the input file path"
	}
	$Count = 1
	try {
		$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$MailAddress = $_.MailAddress.Trim()
			$MemberShip = $_.Membership.Trim()
			Write-Progress -Activity "Prcessed user count : $($Count)" -Status "Currently Processing: $($MailAddress)" -PercentComplete 100
			$Count++
			try {
				$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
				if($MemberShip -eq "owner"){
					Remove-MgGroupOwnerByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				elseif ($MemberShip -eq "member") {
					Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove $($MemberShip) from group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove $($MemberShip) from group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Remove users from list of groups" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

function Remove_BulkUsersFromListOfGroups{
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$UPN = $_.UPN.Trim()
			$MailAddress = $_.MailAddress.Trim()
			$MemberShip = $_.Membership.Trim()
			Write-Progress -Activity "Processed user count: $($Count)"-Status "Currently Processing: $($UPN)" -PercentComplete 100
			$Count++
			try{
				$UserId = (Get-MgUser -UserId $UPN -ErrorAction Stop).Id
				$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
				if($MemberShip -eq "owner"){
					Remove-MgGroupOwnerByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				elseif ($MemberShip -eq "member") {
					Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
				}
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove $($MemberShip) from group" -Value $UPN -Status "Success" -Err " -"
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove $($MemberShip) from group" -Value $UPN -Status "Failed" -Err $_.Exception
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Remove users from list of groups" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}
function AssignLicense {
	param (
		[string]$MailAddress,
		[string[]]$LicenseNames,
		[hashtable]$SkuIDMap
	)
	try {
		$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
		$AssignedLicense = @()
		foreach($LicenseName in $LicenseNames){
			$LicenseId = $SkuIDMap[$FriendlyNameHash[$LicenseName]]
			try {
				$licenseAssignment = @{
					addLicenses = @(@{ skuId = $LicenseId })
					removeLicenses = @()
				}
				Set-MgGroupLicense -GroupId $Group.Id -BodyParameter $licenseAssignment -ErrorAction Stop | Out-Null
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Assign license to group" -Value $LicenseName -Status "Success" -Err " -"
				$AssignedLicense += $LicenseName	
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Assign license to group" -Value $LicenseName -Status "Failed" -Err $_.Exception
			}
		}
		if ([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			if($AssignedLicense -eq 0){
				Write-Host "`nFailed to assign the license $($LicenseNames -join ",") to the group $($MailAddress). Error details are available in the log file." -ForegroundColor Red
			}
			else{
				if($AssignedLicense.Count -ne $LicenseNames.Count){
					$LicenseNames = $AssignedLicense
				}
				Write-Host "`nSuccessfully assigned the license $($LicenseNames -join ",") to the group $($Group.DisplayName)." -ForegroundColor Green
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Assign license to group" -Value $($LicenseNames -split ",") -Status "Failed" -Err $_.Exception
	}
}
function Assign_LicensestoGroup {
	$MailAddress = (Read-Host "`nEnter the group mail address").Trim()
	$LicenseNames = (Read-Host "Enter the license names (comma-separated)") -split "," | ForEach-Object Trim | Where-Object Length
	try {
		$SkuIDMap = @{}
		Get-MgSubscribedSku  | ForEach-Object{
			$SkuIDMap[$_.SkuPartNumber] = $_.SkuId
		}
		AssignLicense -MailAddress $MailAddress -LicenseNames $LicenseNames -SkuIDMap $SkuIDMap
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Assign license to group" -Value " -" -Status "Failed" -Err $_.Exception
	}
}
function Assign_LicensestoGroups {
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$SkuIDMap = @{}
		Get-MgSubscribedSku  | ForEach-Object{
			$SkuIDMap[$_.SkuPartNumber] = $_.SkuId
		}
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$MailAddress = $_.MailAddress.Trim()
			$LicenseName = $_.LicenseName.Trim()
			Write-Progress -Activity "Processed group count: $($Count)" -Status "Currently Processing $($MailAddress)" -PercentComplete 100
			$Count++
			AssignLicense -MailAddress $MailAddress -LicenseNames $LicenseName -SkuIDMap $SkuIDMap
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Assign license to group" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}
function RemoveLicense {
	param (
		[string]$MailAddress,
		[string[]]$LicenseNames,
		[hashtable]$SkuIDMap
	)
	try {
		$Group = Get-MgGroup -Filter "mail eq '$($MailAddress)'"
		$AssignedLicense = @()
		foreach($LicenseName in $LicenseNames){
			$LicenseId = $SkuIDMap[$FriendlyNameHash[$LicenseName]]
			try {
				$licenseAssignment = @{
					addLicenses = @()
					removeLicenses = @($LicenseId) 
				}
				Set-MgGroupLicense -GroupId $Group.Id -BodyParameter $licenseAssignment -ErrorAction Stop | Out-Null
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove license from group" -Value $LicenseName -Status "Success" -Err " -"
				$AssignedLicense += $LicenseName	
			}
			catch {
				Write-Log -EventTime (Get-Date) -GroupName $Group.DisplayName -Action "Remove license from group" -Value $LicenseName -Status "Failed" -Err $_.Exception
			}
		}
		if ([string]::IsNullOrWhiteSpace($InputCsvFilePath)){
			if($AssignedLicense -eq 0){
				Write-Host "`nFailed to remove the license $($LicenseNames -join ",") from the group $($MailAddress). Error details are available in the log file." -ForegroundColor Red
			}
			else{
				if($AssignedLicense.Count -ne $LicenseNames.Count){
					$LicenseNames = $AssignedLicense
				}
				Write-Host "`nSuccessfully removed the license $($LicenseNames -join ",") from the group $($Group.DisplayName)." -ForegroundColor Green
			}
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Remove license from group" -Value $($LicenseNames -split ",") -Status "Failed" -Err $_.Exception
	}
}
function Remove_LicensestoGroup {
	$MailAddress = (Read-Host "`nEnter the group mail address").Trim()
	$LicenseNames = (Read-Host "Enter the license names (comma-separated)") -split "," | ForEach-Object Trim | Where-Object Length
	try {
		$SkuIDMap = @{}
		Get-MgSubscribedSku  | ForEach-Object{
			$SkuIDMap[$_.SkuPartNumber] = $_.SkuId
		}
		RemoveLicense -MailAddress $MailAddress -LicenseNames $LicenseNames -SkuIDMap $SkuIDMap
	}
	catch {
		Write-Log -EventTime (Get-Date) -GroupName $MailAddress -Action "Remove license from group" -Value " -" -Status "Failed" -Err $_.Exception
	}
}
Function Remove_LicensestoGroups{
	if($InputCsvFilePath -eq ""){
		$InputCsvFilePath = Read-Host "`nPlease enter the input file path"
	}
	try {
		$SkuIDMap = @{}
		Get-MgSubscribedSku  | ForEach-Object{
			$SkuIDMap[$_.SkuPartNumber] = $_.SkuId
		}
		$Count = 1
		Import-Csv -Path $InputCsvFilePath | ForEach-Object{
			$MailAddress = $_.MailAddress.Trim()
			$LicenseName = $_.LicenseName.Trim()
			Write-Progress -Activity "Procesed group count: $($Count)" -Status "Currently Processing $($MailAddress)" -PercentComplete 100
			$Count++
			RemoveLicense -MailAddress $MailAddress -LicenseNames $LicenseName -SkuIDMap $SkuIDMap
		}
	}
	catch {
		Write-Log -EventTime (Get-Date) -Action "Remove license from  group" -Value " -" -Status "Failed" -Err $_.Exception
	}
	finally{
		Write-Progress -Activity "Completed" -Completed
	}
}

Function Write-Log{
	param(
		[string]$EventTime,
		[string]$GroupName,
		[string]$Action,
		[string]$Value,
		[string]$Status,
		[string]$Err
	)
	$LogData = [PSCustomObject]@{ 'Event Time' = $EventTime;'Group Name' = $GroupName; 'Action' = $Action; 'Value' = $Value; 'Status' = $Status; 'Error' = $Err;}
	$LogData| Export-Csv -Path $LogFilePath -NoTypeInformation -Append
}

Function Open_OutputFile{
	param(
		[PSCustomObject]$LogFilePath
	)
	Write-Host "`nScript executed successfully..!" -ForegroundColor Green
    if(Test-Path $LogFilePath){
		Write-Host "`nThe log file is available in:" -NoNewline -ForegroundColor Yellow; Write-Host (Resolve-Path $LogFilePath).Path
	}
	Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
	Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to access 3,000+ reports and 450+ management actions across your Microsoft 365 environment. ~~" -ForegroundColor Green `n`n

	if((Test-Path -Path $LogFilePath) -eq "True")
	{
		$Prompt = New-Object -ComObject wscript.shell
		$UserInput = $Prompt.popup("Do you want to open Log file?",` 0,"Open Log File",4)
		if ($UserInput -eq 6)
		{  
			Invoke-Item "$LogFilePath"
		}
	}
}

Connect_MgGraph
$LogFilePath = ".\M365Group_Management_Log_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv"
$FriendlyNameHash = @{}
Import-Csv -Path ".\LicenseFriendlyName.csv" -ErrorAction Stop | ForEach-Object{
    $FriendlyNameHash[$_.Product_Display_Name] = $_.String_Id
}

do {
	if($Action -eq ""){
		Show-Menu
		$GetAction = Read-Host "  Please choose the option"
	}
	else{
		$GetAction = $Action
	}

	switch ($GetAction) {
		 1 { Create_Single_Group }
		 2 { Add_Team_To_Single_Group }
		 3 { Assign_LicensestoGroup }
		 4 { Remove_LicensestoGroup }
		 5 { Remove_UserFromAllGroups }
		 6 { Delete_Single_Group }
		 7 { Restore_Single_Group }
		 8 { Create_Bulk_Groups }
		 9 { Add_UserToBulkGroups }
		10 { Add_BulkUsersToSingleGroup }
		11 { Add_BulkUsersToBulkGroups }
		12 { Add_Team_To_Bulk_Group }
		13 { Assign_LicensestoGroups }
		14 { Remove_LicensestoGroups }
		15 { Remove_UsersFromAllGroups }
		16 { Remove_UserFromListOfGroups }
		17 { Remove_BulkUsersFromListOfGroups }
		18 { Delete_Bulk_Groups }
		19 { Restore_Bulk_Groups }
		 0 {
			Disconnect-Graph | Out-Null
			Open_OutputFile -LogFilePath $LogFilePath
			Exit
		}
		default {
			Write-Host "`nInvalid choice...!" -ForegroundColor Red
		}
	}
    if($MultipleActionsMode.ispresent)
    {                          
    	Start-Sleep -Seconds 2
		$InputCsvFilePath = ""
		$Action = ""
    } 
    else
    {
		Disconnect-Graph | Out-Null
		Open_OutputFile -LogFilePath $LogFilePath
     	Exit
    }
} while ( $MultipleActionsMode.ispresent )