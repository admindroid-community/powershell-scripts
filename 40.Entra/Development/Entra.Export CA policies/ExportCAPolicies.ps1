<#
=============================================================================================
Name:          Export Conditional Access Policies to Excel using PowerShell  
Description:   The script exports all Conditional Access policies to an Excel file. 
Version:       2.2
Website:       o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~  
1. The script generates 6 reports with 33 attributes for detailed CA policy analysis. 
2. The script exports all Conditional Access policies by default. 
3. It generates report on active CA policies. 
4. Finds all disabled CA policies. 
5. It also lists report-only mode CA policies.  
6. Identifies the recently created CA policies for review. 
7. Lists recently modified CA policies for tracking changes. 
8. The script can be executed with MFA-enabled accounts. 
9. It exports reports to CSV format. 
10. The script automatically installs the required Microsoft Graph Beta PowerShell module upon user confirmation. 
11. Supports certificate-based authentication for secure access. 
12. Includes scheduler-friendly functionality for automated reporting. 

For detailed Script execution: https://o365reports.com/2024/02/20/export-conditional-access-policies-to-excel-using-powershell


Change Log:
~~~~~~~~~~
  V1.0 (Feb 20, 2024) - File created
  V2.0 (Feb 27, 2024) - Some CA policies doesn't have creation time. Error handling added for those CA policies.
  V2.1 (Aug 09, 2024) - Error handling added when the script tries to convert a deleted object ID into a name.
  V2.2 (Apr 01, 2025) - Fixed error while converting directory objects to names.
============================================================================================
#>


Param
(
    [switch]$ActiveCAPoliciesOnly,
    [switch]$DisabledCAPoliciesOnly,
    [switch]$ReportOnlyMode,
    [int]$RecentlyCreatedCAPolicies,
    [int]$RecentlyModifiedCAPolicies,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)


$global:DirectoryObjsHash = @{}
$global:ServicePrincipalsHash=@{}
$global:NamedLocationHash=@{}
Function Connect_MgGraph
{
 $MsGraphBetaModule =  Get-Module Microsoft.Graph.Beta -ListAvailable
 if($MsGraphBetaModule -eq $null)
 { 
    Write-host "Important: Microsoft Graph Beta module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
    $confirm = Read-Host Are you sure you want to install Microsoft Graph Beta module? [Y] Yes [N] No  
    if($confirm -match "[yY]") 
    { 
        Write-host "Installing Microsoft Graph Beta module..."
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -AllowClobber
        Write-host "Microsoft Graph Beta module is installed in the machine successfully" -ForegroundColor Magenta 
    } 
    else
    { 
        Write-host "Exiting. `nNote: Microsoft Graph Beta module must be available in your system to run the script" -ForegroundColor Red
        Exit 
    } 
 }
 #Disconnect Existing MgGraph session
 if($CreateSession.IsPresent)
 {
  Disconnect-MgGraph
 }
 #Connecting to MgGraph beta
 Write-Host Connecting to Microsoft Graph...
 if(($TenantId -ne "") -and ($ClientId -ne "") -and ($CertificateThumbprint -ne ""))  
 {  
  Connect-MgGraph  -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint 
 }
 else
 {
  Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All'  -NoWelcome
 }
}
Connect_MgGraph




Function ConvertTo-Name {
    param(
        [Parameter(Mandatory=$true)]
        [Array]$InputIds
    )
    $ConvertedNames = @()
    
    # Process each value in the array
    foreach ($Id in $InputIds) {
         # Check Id-Name pair already exist in hash table
        if($DirectoryObjsHash.ContainsKey($Id))
        {
         $Name=$DirectoryObjsHash[$Id]
         $ConvertedNames += $Name
        }
        # Retrieve the display name for the directory object with the given ID
        else{
         try
         {
          $Name = ((Get-MgBetaDirectoryObject -DirectoryObjectId $Id ).AdditionalProperties["displayName"] )
          if($Name -ne $null)
          {
           $DirectoryObjsHash[$Id]=$Name
           $ConvertedNames += $Name
           
          }
         }
         catch
         {
          Write-Host "Deleted object configured in the CA policy $CAName" -ForegroundColor Red
          Write-Host "Processing CA policies..."
         }
        }        
    } 
     return $ConvertedNames
}

Function ConvertAppIdTo-Name {
    param(
        [Parameter(Mandatory=$true)]
        [Array]$InputIds
    )
    $ConvertedNames = @()
    # Process each value in the array
    foreach ($Id in $InputIds) {
         # Check Id-Name pair already exist in hash table
        if($ServicePrincipalsHash.ContainsKey($Id))
        {
         $Name=$ServicePrincipalsHash[$Id].DisplayName
        }
        else
        { $Name=$Id }
        $ConvertedNames += $Name     
    }
     return $ConvertedNames
}

Function ConvertLocationIdTo-Name {
    param(
        [Parameter(Mandatory=$true)]
        [Array]$InputIds
    )
    $ConvertedNames = @()
    # Process each value in the array
    foreach ($Id in $InputIds) {
         # Check Id-Name pair already exist in hash table
        if($NamedLocationHash.ContainsKey($Id))
        {
         $Name=$NamedLocationHash[$Id].DisplayName
        }
        else
        { $Name=$Id }
        $ConvertedNames += $Name     
    }
     return $ConvertedNames
}

#Prep
$Location=Get-Location
$ExportCSV="$Location\CA_Policies_Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$Result=""   
$Results=@()  
$ProcessedCount=0
$OutputCount=0
#Get all service principals
Write-Progress -Activity "`n     Retrieving service principals..."
$ServicePrincipalsHash=Get-MgBetaServicePrincipal -All | Group-Object -Property AppId -AsHashTable
Write-Progress -Activity "`n     Retrieving named location..."
$NamedLocationHash=$namedLocations = Get-MgBetaIdentityConditionalAccessNamedLocation -All | Group-Object -Property Id -AsHashTable
Write-Host "Exporting CA policies report..." -ForegroundColor Cyan


#Processing all CA polcies
Get-MgBetaIdentityConditionalAccessPolicy -All | Foreach {
 $ProcessedCount++
 $CAName=$_.DisplayName
 $Description=$_.Description
 $CreationTime=$_.CreatedDateTime
 $LastModifiedTime=$_.ModifiedDateTime
 $State=$_.State
 Write-Progress -Activity "`n     Processed CA policies count: $ProcessedCount "`n"  Currently Processing: $CAName"

 #Filter CA policies based on their State
 if($ActiveCAPoliciesOnly.IsPresent -and $State -ne "Enabled")
 {
  return
 }
 elseif($DisabledCAPoliciesOnly.IsPresent -and $State -ne "Disabled" )
 {
  return
 }
 elseif($ReportOnlyMode.IsPresent -and $State -ne "EnabledForReportingButNotEnforced")
 {
  return
 }

 #Calculating recently created and modified days
 if($CreationTime -eq $null)
 {
  $CreationTime = "-"
 }
 else
 {
  $CreatedInDays = (New-TimeSpan -Start $CreationTime).Days
 }

 if($LastModifiedTime -eq $null)
 {
  $LastModifiedTime = "-"
 }
 else
 {
  $ModifiedInDays = (New-TimeSpan -Start $LastModifiedTime).Days
 }

 #Filter for recently created CA policies
 if(($RecentlyCreatedCAPolicies -ne "") -and (($RecentlyCreatedCAPolicies -lt $CreatedInDays) -or ($CreationTime -eq "-")))
 {
  return
 }
 
 #Filter for recently modified CA polcies
 if(($RecentlyModifiedCAPolicies -ne "") -and (($RecentlyModifiedCAPolicies -lt $ModifiedInDays) -or ($LastModifiedTime -eq "-") ))
 {
  return
 }


 #Assignments
 $Conditions=$_.Conditions
 $IncludeUsers=$Conditions.Users.IncludeUsers
 $ExcludeUsers=$Conditions.Users.ExcludeUsers
 $IncludeGroups=$Conditions.Users.IncludeGroups
 $ExcludeGroups=$Conditions.Users.ExcludeGroups
 $IncludeRoles=$Conditions.Users.IncludeRoles
 $ExcludeRoles=$Conditions.Users.ExcludeRoles
 $IncludeGuestsOrExtUsers=$Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes
 $ExcludeGuestsOrExtUsers=$Conditions.Users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes

 #Convert id to names for Assignment properties
 if($IncludeUsers.Count -ne 0 -and ($IncludeUsers -ne 'All' -and $IncludeUsers -ne 'None' ))
 { 
  $IncludeUsers=ConvertTo-Name -InputIds $IncludeUsers
 }
 $IncludeUsers=$IncludeUsers -join ","
 
 if(($ExcludeUsers.Count -ne 0) -and ($ExcludeUsers -ne 'GuestsOrExternalUsers'  ))
 {
  $ExcludeUsers=ConvertTo-Name -InputIds $ExcludeUsers
 }
 $ExcludeUsers=$ExcludeUsers -join ","
 if($IncludeGroups.Count -ne 0) 
 {
  $IncludeGroups=ConvertTo-Name -InputIds $IncludeGroups
 }
 $IncludeGroups=$IncludeGroups -join ","
 if($ExcludeGroups.Count -ne 0) 
 {
  $ExcludeGroups=ConvertTo-Name -InputIds $ExcludeGroups
 }
 $ExcludeGroups=$ExcludeGroups -join ","
 if($IncludeRoles.Count -ne 0 -and ($IncludeRoles -ne 'All' -and $IncludeRoles -ne 'None' ))
 {
  $IncludeRoles=ConvertTo-Name -InputIds $IncludeRoles
 }
 $IncludeRoles=$IncludeRoles -join ","
 if($ExcludeRoles.Count -ne 0) 
 {
  $ExcludeRoles=ConvertTo-Name -InputIds $ExcludeRoles
 }
 $ExcludeRoles=$ExcludeRoles -join ","
 
 $IncludeGuestsOrExtUsers=$IncludeGuestsOrExtUsers -join ","
 $ExcludeGuestsOrExtUsers=$ExcludeGuestsOrExtUsers -join ","
 


 #Target Resources
 $IncludeApplications=$_.Conditions.Applications.IncludeApplications
 $ExcludeApplications=$_.Conditions.Applications.ExcludeApplications
 $AuthContext=$_.Conditions.Applications.IncludeAuthenticationContextClassReferences
 $UserAction=$_.Conditions.Applications.IncludeUserActions
 $UserAction=$UserAction -join ","
 
 #Convert id to names for Target resource properties
 if($IncludeApplications.Count -ne 0 -and ($IncludeApplications -ne 'All' -and $IncludeApplications -ne 'None' ))
 {
  $IncludeApplications=ConvertAppIdTo-Name -InputIds $IncludeApplications
 }
 $IncludeApplications=$IncludeApplications -join ","
 if($ExcludeApplications.Count -ne 0)
 {
  $ExcludeApplications=ConvertAppIdTo-Name -InputIds $ExcludeApplications
 }
 $ExcludeApplications=$ExcludeApplications -join ","



 #Conditions
 $UserRisk=$_.Conditions.UserRiskLevels
 $SigninRisk=$_.Conditions.SignInRiskLevels
 $ClientApps=$_.Conditions.ClientAppTypes
 $IncludeDevicePlatform=$_.Conditions.Platforms.IncludePlatforms
 $ExcludeDevicePlatform=$_.Conditions.Platforms.ExcludePlatforms
 $IncludeLocations=$_.Conditions.Locations.IncludeLocations
 $ExcludeLocations=$_.Conditions.Locations.ExcludeLocations

 $UserRisk=$UserRisk -join ","
 $SigninRisk=$SigninRisk -join ","
 $ClientApps=$ClientApps -join ","
 $IncludeDevicePlatform=$IncludeDevicePlatform -join ","
 $ExcludeDevicePlatform=$ExcludeDevicePlatform -join ","

 #Convert location id to Name
 if($IncludeLocations.Count -ne 0 -and $IncludeLocations -ne 'All' -and $IncludeLocations -ne 'AllTrusted')
 {
  $IncludeLocations=ConvertLocationIdTo-Name -InputIds $IncludeLocations
 }
 $IncludeLocations=$IncludeLocations -join ","

 if($ExcludeLocations.Count -ne 0)
 {
  $ExcludeLocations=ConvertLocationIdTo-Name -InputIds $ExcludeLocations
 }
 $ExcludeLocations=$ExcludeLocations -join ","



 #Grant Control
 $AccessControl=$_.GrantControls.BuiltInControls -join ","
 $AccessControlOperator=$_.GrantControls.Operator
 $AuthenticationStrength=$_.GrantControls.AuthenticationStrength.DisplayName
 $AuthenticationStrengthAllowedCombo=$_.GrantControls.AuthenticationStrength.AllowedCombinations -join ","

 #Session Control
 $AppEnforcedRestrictions=$_.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
 $CloudAppSecurity=$_.SessionControls.CloudAppSecurity.IsEnabled
 $CAEMode=$_.SessionControls.ContinuousAccessEvaluation.Mode
 $DisableResilienceDefaults=$_.SessionControls.DisableResilienceDefaults
 $PersistentBrowser=$_.SessionControls.PersistentBrowser.Mode
 $IsSigninFrequencyEnabled=$_.SessionControls.SignInFrequency.IsEnabled
 $SignInFrequencyValue="$($_.SessionControls.SignInFrequency.Value) $($_.SessionControls.SignInFrequency.Type)"
 
 

 $OutputCount++
  $Result=@{'CA Policy Name'=$CAName;
            'Description'=$Description;
            'Creation Time'=$CreationTime;
            'Modified Time'=$LastModifiedTime;
            'Include Users'=$IncludeUsers;
            'Exclude Users'=$ExcludeUsers;
            'Include Groups'=$IncludeGroups;
            'Exclude Groups'=$ExcludeGroups;
            'Include Roles'=$IncludeRoles;
            'Exclude Roles'=$ExcludeRoles;
            'Include Guests or Ext Users'=$IncludeGuestsOrExtUsers;
            'Exclude Guests or Ext Users'=$ExcludeGuestsOrExtUsers;
            'Include Applications'=$IncludeApplications;
            'Exclude Applications'=$ExcludeApplications;
            'User Action'=$UserAction;
            'User Risk'=$UserRisk;
            'Signin Risk'=$SigninRisk;
            'Client Apps'=$ClientApps;
            'Include Device Platform'=$IncludeDevicePlatform;
            'Exclude Device Platform'=$ExcludeDevicePlatform;
            'Include Locations'=$IncludeLocations;
            'Exclude Locations'=$ExcludeLocations;
            'Access Control'=$AccessControl;
            'Access Control Operator'=$AccessControlOperator;
            'Authentication Strength'=$AuthenticationStrength;
            'Auth Strength Allowed Combo'=$AuthenticationStrengthAllowedCombo;
            'App Enforced Restrictions Enabled'=$AppEnforcedRestrictions;
            'Cloud App Security'=$CloudAppSecurity;
            'CAE Mode'=$CAEMode;
            'Disable Resilience Defaults'=$DisableResilienceDefaults;
            'Is Signin Frequency Enabled'=$IsSigninFrequencyEnabled;
            'Signin Frequency Value'=$SignInFrequencyValue;
            'State'=$State}
  $Results= New-Object PSObject -Property $Result  
  $Results | Select-Object 'CA Policy Name','Description','Creation Time','Modified Time','State','Include Users','Exclude Users','Include Groups','Exclude Groups','Include Roles','Exclude Roles',
  'Include Guests or Ext Users','Exclude Guests or Ext Users','Include Applications','Exclude Applications','User Action','User Risk','Signin Risk','Client Apps','Include Device Platform',
  'Exclude Device Platform','Include Locations','Exclude Locations','Access Control','Access Control Operator','Authentication Strength','Auth Strength Allowed Combo',
  'App Enforced Restrictions Enabled','Cloud App Security','CAE Mode','Disable Resilience Defaults','Is Signin Frequency Enabled','Signin Frequency Value'| Export-Csv -Path $ExportCSV -Notype -Append 
 }


#Open output file after execution 
 If($OutputCount -eq 0)
 {
  Write-Host No data found for the given criteria
 }
 else
 {
  Write-Host `nThe output file contains $OutputCount CA policies.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {
   Write-Host `nThe Output file available in:  -NoNewline -ForegroundColor Yellow
   Write-Host $ExportCSV 
   Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
   Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
 
   $Prompt = New-Object -ComObject wscript.shell      
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
   {   
    Invoke-Item "$ExportCSV"   
   } 
  }
 }
