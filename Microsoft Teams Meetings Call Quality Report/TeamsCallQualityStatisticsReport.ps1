<#
=============================================================================================
Name: Export Microsoft Teams Call Quality Reports Using PowerShell
Version:1.0
Website:m365scripts.com

~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Generate 6 Microsoft Teams Call quality Reports by default. 
2. Exports Teams call quality report results as a CSV file. 
3. The script can be executed with an MFA-enabled account too. 
4. The script uses CQD PowerShell and installs CQD PowerShell (if not installed already) upon your confirmation. 
5. The script is scheduler friendly. 
6. The script sets for 28 days by default as EUII data only available for 28 days. 

For detailed script execution:https://m365scripts.com/microsoft-teams/check-teams-meeting-call-quality-using-powershell/

============================================================================================
#>
param (
    [Parameter(Mandatory=$false)]
    [ValidateSet('MeetingInfo','NetworkQuality','AudioHealth','VideoHealth','ScreenSharingHealth','DeviceUsed')]
    [string[]]$ReportRequired
)

#CQDPowerShell Module check
$CQDModule = Get-Module CQDPowerShell -ListAvailable
if($CQDModule.count -eq 0) {
    Write-Host CQDPowerShell module is not available. -ForegroundColor Yellow
    $confirm = Read-Host Do install want to install module? [Y] yes  [N] No
    if($confirm -match "[Yy]") {
        Write-Host "Installing CQDPowerShell module... "
        Install-Module -Name CQDPowerShell -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        Import-Module CQDPowerShell
    } 
    else {
        Write-Host "CQDPowerShell is required. To install use 'Install-Module CQDPowerShell' cmdlet." -ForegroundColor Red
        Exit
    }
}

#Retriving data for past 28 days
$EndDate = Get-Date
$StartDate = (Get-Date).AddDays(-27)

#get file directory
$OutputPath = $PSScriptRoot

#Initialize Report required 
if($ReportRequired.count -eq 0)
{
    $ReportRequired = 'MeetingInfo','NetworkQuality','AudioHealth','VideoHealth','ScreenSharingHealth','DeviceUsed'
}

Write-Host `nRetrieving Teams meeting call quality statistics from $StartDate to $EndDate... -ForegroundColor Yellow

foreach($report in $ReportRequired) {
    switch($report) {
        #Meeting info
        "MeetingInfo" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_MeetingInfo_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $ExportResult = ""   
            $ExportResults = @()
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id","AllStreams.Organizer UPN","AllStreams.First UPN","AllStreams.Second UPN","AllStreams.Transport","AllStreams.Session Type","AllStreams.Start Time","AllStreams.End Time"
            $measures = "Measures.Avg Network Score","Measures.Total Audio Stream Duration (Minutes)"
            Write-host "Call quality statistics Report - Meeting info " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object {
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $OrganizerUPN = $data.'Organizer UPN'
                    $FirstUPN = $data.'First UPN'
                    $SecondUPN = $data.'Second UPN'
                    $Transport = $data.'Transport'
                    $SessionType = $data.'Session Type'
                    $AvgNetworkScore = $data.'Avg Network Score'
                    $JoinTime = $data.'Start Time'
                    $EndTime = $data.'End Time'
                    if(!([string]::IsNullOrEmpty($JoinTime))) {
                        $EventStartTime = (Get-date($JoinTime)).ToLocalTime()
                    }
                    else {
                        $EventStartTime = "-"
                    }
                    if(!([string]::IsNullOrEmpty($EndTime))) {
                        $EventEndTime = (Get-date($EndTime)).ToLocalTime()
                    }
                    else {
                        $EventEndTime = "-"
                    }
                    if((!([string]::IsNullOrEmpty($JoinTime))) -and (!([string]::IsNullOrEmpty($EndTime))))
                    {
                        $Duration = $EventStartTime - $EventEndTime
                        $DurationinSeconds = ($Duration).TotalSeconds
                        $TimeSpanDuration =  [timespan]::fromseconds($DurationinSeconds)
                        $AttendedDuration = ("{0:hh\:mm\:ss}" -f $TimeSpanDuration)
                    }
                    else {
                        $AttendedDuration = "-"
                    }
                    
                    if($SessionType -eq "P2P") {
                        $SessionType = "Peer-to-Peer"
                    }
                    elseif($SessionType -eq "Conf") {
                        $SessionType = "Meeting"
                    }
                    else {
                        $SessionType = "-"
                    }
                    # Write-Host New data `n $data
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"Organizer UPN" = $OrganizerUPN;"UPN of first end of stream" = $FirstUPN ;"UPN of second end of stream" = $SecondUPN;"Network Transport Type" = $Transport;"Session Type" = $SessionType;"Avg Network Score (0-10)" = $AvgNetworkScore;"Total Stream Duration" = $AttendedDuration;"Start time" = $EventStartTime;"End time" = $EventEndTime}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","Organizer UPN","UPN of first end of stream","UPN of second end of stream","Session Type","Network Transport Type","Avg Network Score (0-10)","Start time","End time","Total Stream Duration" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append
                }
            }
            break
        }

        #Network quality
        "NetworkQuality" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_NetworkQuality_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id","AllStreams.First UPN","AllStreams.Second UPN","AllStreams.Session Type","AllStreams.First ASN ISP Name","AllStreams.Second ASN ISP Name"
            $measures = "Measures.Avg Jitter","Measures.Avg Packet Loss Rate","Measures.Avg Round Trip"
            Write-host "Call quality statistics Report - Network Quality " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object {
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $FirstUPN = $data.'First UPN'
                    $SecondUPN = $data.'Second UPN'
                    $SessionType = $data.'Session Type'
                    $FirstISP = $data.'First ASN ISP Name'
                    $SecondISP = $data.'Second ASN ISP Name'
                    $AvgJitter = $data.'Avg Jitter'
                    $AvgPacketLoss = $data.'Avg Packet Loss Rate'
                    $AvgRoundTrip = $data.'Avg Round Trip'
                    if($SessionType -eq "P2P") {
                        $SessionType = "Peer-to-Peer"
                    }
                    elseif($SessionType -eq "Conf") {
                        $SessionType = "Meeting"
                    }
                    else {
                        $SessionType = "-"
                    }
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"UPN of first end of stream" = $FirstUPN ;"UPN of second end of stream" = $SecondUPN;"Session Type" = $SessionType;"ISP for first end of stream" = $FirstISP;"ISP for second end of stream" = $SecondISP;"Average Jitter (ms)" = $AvgJitter;"Average Packet Loss Rate (0-1)" = $AvgPacketLoss;"Average Round Trip (ms)" = $AvgRoundTrip}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","UPN of first end of stream","UPN of second end of stream","Session Type","ISP for first end of stream","ISP for second end of stream","Average Jitter (ms)","Average Packet Loss Rate (0-1)","Average Round Trip (ms)" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append

                }
            }
            break
        }

        #Audio health
        "AudioHealth" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_AudioHealth_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id"
            $measures = "Measures.Audio Good Call Stream Count","Measures.Audio Poor Call Stream Count","Measures.Audio Unclassified Call Stream Count","Measures.Audio Poor Percentage","Measures.Audio Stream Count"
            Write-host "Call quality statistics Report - Audio Health " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object {
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $AudioGoodCall = $data.'Audio Good Call Stream Count'
                    $AudioPoorCall = $data.'Audio Poor Call Stream Count'
                    $AudioUnclassifiedCall = $data.'Audio Unclassified Call Stream Count'
                    $AudioPoorPercentage = $data.'Audio Poor Percentage'
                    $TotalAudioStream = $data.'Audio Stream Count'
                    if([string]::IsNullOrEmpty($AudioPoorPercentage)) {
                        $AudioPoorPercentage = "-"
                    }
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"Good Call Stream" = $AudioGoodCall;"Poor Call Stream" = $AudioPoorCall;"Unclassified Call Stream" = $AudioUnclassifiedCall;"Poor Percentage" = $AudioPoorPercentage;"Total Audio Stream Count" = $TotalAudioStream}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","Good Call Stream","Poor Call Stream","Unclassified Call Stream","Poor Percentage","Total Audio Stream Count" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append
                }
            }           
            break
        }

        #Video health
        "VideoHealth" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_VideoHealth_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id"
            $measures = "Measures.Video Good Stream Count","Measures.Video Poor Stream Count","Measures.Video Unclassified Stream Count","Measures.Video Poor Percentage","Measures.Video Stream Count"
            Write-host "Call quality statistics Report - Video Health " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object {
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $VideoGoodCall = $data.'Video Good Stream Count'
                    $VideoPoorCall = $data.'Video Poor Stream Count'
                    $VideoUnclassifiedCall = $data.'Video Unclassified Stream Count'
                    $VideoPoorPercentage = $data.'Video Poor Percentage'
                    $TotalVideoStream = $data.'Video Stream Count'
                    if([string]::IsNullOrEmpty($VideoPoorPercentage)) {
                        $VideoPoorPercentage = "-"
                    }
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"Good Call Stream" = $VideoGoodCall;"Poor Call Stream" = $VideoPoorCall;"Unclassified Call Stream" = $VideoUnclassifiedCall;"Poor Percentage" = $VideoPoorPercentage;"Total Video Stream Count" = $TotalVideoStream}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","Good Call Stream","Poor Call Stream","Unclassified Call Stream","Poor Percentage","Total Video Stream Count" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append
                }
            }
            break
        }

        #ScreenSharing health 
        "ScreenSharingHealth" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_ScreenSharingHealth_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id"
            $measures = "Measures.VBSS Good Stream Count","Measures.VBSS Poor Stream Count","Measures.VBSS Unclassified Stream Count","Measures.VBSS Poor Percentage","Measures.VBSS Stream Count"
            Write-host "Call quality statistics Report - ScreenSharing Health " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object {
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $VBSSGoodCall = $data.'VBSS Good Stream Count'
                    $VBSSPoorCall = $data.'VBSS Poor Stream Count'
                    $VBSSUnclassifiedCall = $data.'VBSS Unclassified Stream Count'
                    $VBSSPoorPercent = $data.'VBSS Poor Percentage'
                    $TotalVBSSStream = $data.'VBSS Stream Count'
                    if([string]::IsNullOrEmpty($VBSSPoorPercent)) {
                        $VBSSPoorPercent = "-"
                    }
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"Good Call Stream" = $VBSSGoodCall;"Poor Call Stream" = $VBSSPoorCall;"Unclassified Call Stream" = $VBSSUnclassifiedCall;"Poor Percentage" = $VBSSPoorPercent;"Total ScreenSharing Stream Count" = $TotalVBSSStream}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","Good Call Stream","Poor Call Stream","Unclassified Call Stream","Poor Percentage","Total ScreenSharing Stream Count" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append
                }
            }
            break
        }

        #Device usage
        "DeviceUsed" {
            $OutfilePath = "$OutputPath/CallQualityStatisticsReport_DeviceUsage_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $dimensions = "AllStreams.Is Teams","AllStreams.Date","AllStreams.Conference Id","AllStreams.First Compute Device Name","AllStreams.Second Compute Device Name"
            $measures = "Measures.Total Audio Stream Duration (Minutes)","Measures.First User Count","Measures.Second User Count"
            Write-host "Call quality statistics Report - Device Usage " -NoNewline
            Get-CQDData -StartDate $StartDate -EndDate $EndDate -Dimensions $dimensions -Measures $measures -OutputType DataTable | ForEach-Object { 
                $data = $_
                $IsTeams = $data.'Is Teams'
                if($IsTeams)
                {
                    $date = $data.'Date'
                    $ConferenceId = $data.'Conference Id'
                    $FirstUserDevice = $data.'First Compute Device Name'
                    $SecondUserDevice = $data.'Second Compute Device Name'
                    $TotalAudioDuration = $data.'Total Audio Stream Duration (Minutes)'
                    $FirstUserDeviceCount = $data.'First User Count'
                    $SecondUserDeviceCount = $data.'Second User Count'
                    $ExportResult = @{"Date" = $date;"Meeting/Call Id" = $ConferenceId;"Device used in first end of stream" = $FirstUserDevice;"Device used in second end of stream" = $SecondUserDevice;"Total Audio Stream Duration (Minutes)" = $TotalAudioDuration;"First end user device used count" = $FirstUserDeviceCount;"Second end user device used count" = $SecondUserDeviceCount}
                    $ExportResults = New-Object PSObject -Property $ExportResult  
                    $ExportResults | Select-Object "Date","Meeting/Call Id","Device used in first end of stream","Device used in second end of stream","Total Audio Stream Duration (Minutes)","First end user device used count","Second end user device used count" | Export-Csv -Path $OutfilePath -NoTypeInformation -Append
                }
            }      
            break
        }

        Default {
            Write-Host "`nInvalid report name!`n" -ForegroundColor Red
            return
        }
    }
}

#show output file's folder location
Write-Host `n "The location of output file is : " -NoNewline -ForegroundColor Yellow; Write-Host "$OutputPath" `n 
Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
