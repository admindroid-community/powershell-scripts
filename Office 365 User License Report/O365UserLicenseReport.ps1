#Using this script administrator can identify all licensed users with their assigned licenses, services, and its status.

Param
(
 [Parameter(Mandatory = $false)]
    [string]$UserNamesFile
)


Function Get_UsersLicenseInfo
{
  $LicensePlanWithEnabledService=""
  $FriendlyNameOfLicensePlanWithService=""
  $upn=$_.userprincipalname
  $Country=$_.Country
  if([string]$Country -eq "")
  {
   $Country="-"
  }
  Write-Progress -Activity "`n     Exported user count:$LicensedUserCount "`n"Currently Processing:$upn"
  #Get all asssigned SKU for current user
  $Skus=$_.licenses.accountSKUId
  $LicenseCount=$skus.count
  $count=0
  #Loop through each SKUid
  foreach($Sku in $Skus)  #License loop
  {
   #Convert Skuid to friendly name
   $LicenseItem= $Sku -Split ":" | Select-Object -Last 1
   $EasyName=$FriendlyNameHash[$LicenseItem]
   if(!($EasyName))
   {$NamePrint=$LicenseItem}
   else
   {$NamePrint=$EasyName}
   #Get all services for current SKUId
   $Services=$_.licenses[$count].ServiceStatus
   if(($Count -gt 0) -and ($count -lt $LicenseCount))
   {
    $LicensePlanWithEnabledService=$LicensePlanWithEnabledService+","
    $FriendlyNameOfLicensePlanWithService=$FriendlyNameOfLicensePlanWithService+","
   }
   $DisabledServiceCount = 0
   $EnabledServiceCount=0
   $serviceExceptDisabled=""
   $FriendlyNameOfServiceExceptDisabled=""
   foreach($Service in $Services) #Service loop
   {
    $flag=0
    $ServiceName=$Service.ServicePlan.ServiceName
    if($service.ProvisioningStatus -eq "Disabled")
    {
     $DisabledServiceCount++
    }
    else
    {
     $EnabledServiceCount++
     if($EnabledServiceCount -ne 1)
     {
      $serviceExceptDisabled =$serviceExceptDisabled+","
     }
     $serviceExceptDisabled =$serviceExceptDisabled+$ServiceName
     $flag=1
    }
    #Convert ServiceName to friendly name
    for($i=0;$i -lt $ServiceArray.length;$i +=2)
    {
     $ServiceFriendlyName = $ServiceName
     $Condition = $ServiceName -Match $ServiceArray[$i]
     if($Condition -eq "True")
     {
      $ServiceFriendlyName=$ServiceArray[$i+1]
      break
     }
    }
    if($flag -eq 1)
    {
     if($EnabledServiceCount -ne 1)
     {
     $FriendlyNameOfServiceExceptDisabled =$FriendlyNameOfServiceExceptDisabled+","
    }
    $FriendlyNameOfServiceExceptDisabled =$FriendlyNameOfServiceExceptDisabled+$ServiceFriendlyName
   }
   #Store Service and its status in Hash table
   $Result = @{'DisplayName'=$_.Displayname;'UserPrinciPalName'=$upn;'LicensePlan'=$Licenseitem;'FriendlyNameofLicensePlan'=$nameprint;'ServiceName'=$service.ServicePlan.ServiceName;
   'FriendlyNameofServiceName'=$serviceFriendlyName;'ProvisioningStatus'=$service.ProvisioningStatus}
    $Results = New-Object PSObject -Property $Result
    $Results |select-object DisplayName,UserPrinciPalName,LicensePlan,FriendlyNameofLicensePlan,ServiceName,FriendlyNameofServiceName,
        ProvisioningStatus | Export-Csv -Path $ExportCSV -Notype -Append
  }
  if($Disabledservicecount -eq 0)
  {
   $serviceExceptDisabled ="All services"
   $FriendlyNameOfServiceExceptDisabled="All services"
  }
  $LicensePlanWithEnabledService=$LicensePlanWithEnabledService + $Licenseitem +"[" +$serviceExceptDisabled +"]"
  $FriendlyNameOfLicensePlanWithService=$FriendlyNameOfLicensePlanWithService+ $NamePrint + "[" + $FriendlyNameOfServiceExceptDisabled +"]"
  #Increment SKUid count
  $count++
 }
 $Output=@{'Displayname'=$_.Displayname;'UserPrincipalName'=$upn;Country=$Country;'LicensePlanWithEnabledService'=$LicensePlanWithEnabledService;
      'FriendlyNameOfLicensePlanAndEnabledService'=$FriendlyNameOfLicensePlanWithService}
 $Outputs= New-Object PSObject -Property $output
 $Outputs | Select-Object Displayname,userprincipalname,Country,LicensePlanWithEnabledService,FriendlyNameOfLicensePlanAndEnabledService | Export-Csv -path $ExportSimpleCSV -NoTypeInformation -Append
}


Function main()
{
 #Clean up session
 Get-PSSession | Remove-PSSession
 #Connect AzureAD from PowerShell
 Connect-MsolService
 #Set output file
 $ExportCSV=".\DetailedO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
 $ExportSimpleCSV=".\SimpleO365UserLicenseReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
 #FriendlyName list for license plan and service
 $FriendlyNameHash=Get-Content -Raw -Path .\LicenseFriendlyName.txt -ErrorAction Stop | ConvertFrom-StringData
 $ServiceArray=Get-Content -Path .\ServiceFriendlyName.txt -ErrorAction Stop
 #Hash table declaration
 $Result=""
 $Results=@()
 $output=""
 $outputs=@()
 #Get licensed user
 $LicensedUserCount=0

 #Check for input file/Get users from input file
 if([string]$UserNamesFile -ne "")
 {
  #We have an input file, read it into memory
  $UserNames=@()
  $UserNames=Import-Csv -Header "DisplayName" $UserNamesFile
  $userNames
  foreach($item in $UserNames)
  {
   Get-MsolUser -UserPrincipalName $item.displayname | where{$_.islicensed -eq "true"} | Foreach{
   Get_UsersLicenseInfo
   $LicensedUserCount++}
  }
 }

 #Get all licensed users
 else
 {
  Get-MsolUser -All | where{$_.islicensed -eq "true"} | Foreach{
  Get_UsersLicenseInfo
  $LicensedUserCount++}
 }


 #Open output file after execution
 Write-Host Detailed report available in: $ExportCSV
 Write-host Simple report available in: $ExportSimpleCSV
 $Prompt = New-Object -ComObject wscript.shell
 $UserInput = $Prompt.popup("Do you want to open output files?",`
 0,"Open Files",4)
 If ($UserInput -eq 6)
 {
  Invoke-Item "$ExportCSV"
  Invoke-Item "$ExportSimpleCSV"
 }
}
 . main
