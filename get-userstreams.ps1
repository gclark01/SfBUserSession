<#
.SYNOPSIS
Test**** Main script to pull back data using Get-CsUserSession into CSV.  The main purpose of the data output will be used in Power BI reporting.

.DESCRIPTION
This scripts calls upon multiple modules to handle credential management, processing of users in batches, managing the pssession to prevent time-out and utilizing
multiple functions to pull data out of the Get-CsUserSesion to be populated in csv. 

.NOTES
This script is currently under development along with the accomplining Power BI Reports.
Any questions / feedback please reach out to Geoff Clark gclark@microsoft.com
#>

# Enabling Advanced Functions
#[cmdletbinding()]
Param(
    [parameter (Mandatory = $true)]
    [ValidateSet("Report")]
    [string]$Type,
    [parameter(Mandatory = $true)]
    [datetime]$startTime,
    [parameter(Mandatory = $true)]
    [datetime]$endTime,
    [string]$CsvUsers,
    [parameter(Mandatory = $true)]
    [string]$ReportPath
)


# Global Variables
$endTime = $endTime.AddDays(1)
$Date = (Get-Date).ToString("MMddyyyy_hhmmss")
$Global:AudioFileName = $Date + "-" + "audio-records.csv"
$Global:AudioReports = $ReportPath + "\" + $Global:AudioFileName
$Global:ReliabilityFileName = $Date + "-" + "reliability-records.csv"
$Global:ReliabilityReports = $ReportPath + "\" + $Global:ReliabilityFileName
$Global:VbssFileName = $Date + "-" + "vbss-records.csv"
$Global:VbssReports = $ReportPath + "\" + $Global:VbssFileName
$Global:VideoFileName = $Date + "-" + "video-records.csv"
$Global:VideoReports = $ReportPath + "\" + $Global:VideoFileName
$Global:AppShareFileName = $Date + "-" + "appshare-records.csv"
$Global:AppShareReports = $ReportPath + "\" + $Global:AppShareFileName
$Global:RMCFileName = $Date + "-" + "rmc-records.csv"
$Global:RMCReports = $ReportPath + "\" + $Global:RMCFileName

# Set path to save reports
if (! ([IO.Directory]::Exists($ReportPath))) {
    try {
        New-Item -Path $ReportPath -ItemType Directory | Out-Null
    }
    catch {
        throw;
    }
}

#Create new File (Report)
New-Item -Path $ReportPath -Name $Global:AudioFileName -ItemType File | Out-Null
New-Item -Path $ReportPath -Name $Global:ReliabilityFileName -ItemType File | Out-Null
New-Item -Path $ReportPath -Name $Global:VbssFileName -ItemType File | Out-Null
New-Item -Path $ReportPath -Name $Global:VideoFileName -ItemType File | Out-Null
New-Item -Path $ReportPath -Name $Global:AppShareFileName -ItemType File | Out-Null
New-Item -Path $ReportPath -Name $Global:RMCFileName -ItemType File | Out-Null



# Remove script module if previously loaded
if (Get-Module -Name PSSessions) {
    Remove-Module -Name PSSessions
}
if (Get-Module -Name CredentialManager) {
    Remove-Module -Name CredentialManager
}
if (Get-Module -Name BatchUsers) {
    Remove-Module -Name BatchUsers
}
if (Get-Module -Name SessionData) {
    Remove-Module -Name SessionData
}

# Load script Modules
# Import-Module LyncOnlineConnector
Import-Module .\Modules\CredentialManager.psd1
Import-Module .\Modules\PSSessions.psm1
Import-Module .\Modules\BatchUsers.psm1
Import-Module .\Modules\SessionData.psm1


# This starts the inital logon process and loads the remote ps session
$scriptStart = Get-Date
Get-MySession

# Get all users
if ($CsvUsers) {
    $OnlineUsers = Import-Csv $CsvUsers
}
else {
    $OnlineUsers = Get-CsOnlineUser -Filter {Enabled -eq $True} -WarningAction SilentlyContinue | Select-Object UserPrincipalName, SipAddress | Sort-Object -Property SipAddress
}

# Process users into batches
$Batch = Get-BatchUsers -Users $OnlineUsers

# For Testing Purpoose
Write-Host "Count of users found $($OnlineUsers.Count)" -BackgroundColor Red


#Process batch of users
foreach ($userKey in $Batch.Values) {
    # Write how many Batches have been created
    Write-Host "Created $($Batch.Count) unique user batches to be processed.  Max batch size is 500 users" -ForegroundColor Yellow

    # Define variables used during processing of user batches
    $batchCount ++
    $pStartTime = Get-Date

    # Write what batch is currently being processed
    Write-Host "Processing $($batchCount) batch of users out of $($Batch.Count) batches created" -ForegroundColor Yellow
    if ($batchCount -gt 1) {
       
        # Determine how much time is left in comparison to how long the last batch took to run
        $Global:TimeRemaining = New-TimeSpan -Start $Global:SessionTime -End $pEndTime

        # Validate how much time is left, before continuing
        if ($Global:TimetoComplete.TotalSeconds -gt $Global:TimeRemaining.TotalSeconds ) {
            # If the batch took longer than expected, we should update the session before moving forward
            Get-MySession -ForceUpdate
        }
        else {
            # Check to see if we are over the session timer limit of 45 minutes
            Get-MySession
        }
    }
    foreach ($userObject in $userKey) {
        switch ($Type) {
            AudioEvents {  
                $AudioEvents = Get-AudioEvents -userHash $userObject -startTime $startTime -endTime $endTime
                
                # Write results to CSV
                if ($AudioEvents) {
                    $AudioEvents | Export-Csv -Path "c:\temp\audioevents.csv" -NoTypeInformation -Append
                }
            }
            AudioQuality {
                $AudioQuality = Get-AudioQuality -userHash $userObject -startTime $startTime -endTime $endTime

                # Write results to CSV
                if ($AudioQuality) {
                    $AudioQuality | Export-Csv -Path "c:\temp\audioquality.csv" -NoTypeInformation -Append
                }
            }
            VideoApplicationSharing {
                $VideoAppSharing = Get-VideoAppSharingStreams -userHash $userObject -startTime $startTime -endTime $endTime
                
                # Write results to CSV
                if ($VideoAppSharing) {
                    $VideoAppSharing | Export-Csv -Path "c:\temp\vbssquality.csv" -NoTypeInformation -Append
                }
            }
            IMFED {
                $IMFED = Get-IMFederatedDomains -userHash $userObject -startTime $startTime -endTime $endTime

                if ($IMFED) {
                    $IMFED | Export-Csv -Path "c:\temp\imfeddomains.csv" -NoTypeInformation -Append
                }
            }
            SetupOrDrop {
                $SoD = Get-SetupOrDrops -userHash $userObject -startTime $startTime -endTime $endTime

                if ($SoD) {
                    $SoD | Export-Csv -Path "c:\temp\reliability.csv" -NoTypeInformation -Append
                }
            }
            RMC {
                $RMC = Get-RMC -userHash $userObject -startTime $startTime -endTime $endTime

                if ($RMC) {
                    $RMC | Export-Csv -Path "c:\temp\ratemycall.csv" -NoTypeInformation -Append
                }
            }
            Report {
                Get-QoEReport -userHash $userObject -startTime $startTime -endTime $endTime

            }
            Default {
                Write-Error -Message "You must select a report type to run!"
            }
        }
        
    }
    
    # Define variables used to determine length of time spent processing the user batches.
    $pEndTime = Get-Date
    $Global:TimetoComplete = New-TimeSpan -Start $pStartTime -End $pEndTime
}

$endScript = Get-Date
$ScriptTimetoComplete = New-TimeSpan -Start $scriptStart -End $endScript
Write-Host "Total Time to complete Script: $($ScriptTimetoComplete)" -ForegroundColor Yellow
