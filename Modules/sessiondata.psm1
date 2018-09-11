function Get-QoEReport {
    param (
        [object]$userHash,
        [datetime]$startTime,
        [datetime]$endTime
    )

    # Get User Sessions
    $sessions = getSession -SipAddress $userHash.SipAddress.Split(":")[1] -startTime $startTime -endTime $endTime
    Write-Host "Processing data for $($userHash.SipAddress.Split(":")[1])"
    Write-Host "Sessions identified $($sessions.count)"

    # Execute Reports
    $AudioRecords = Get-AudioRecords -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $ReliabilityRecords = Get-SetupOrDrops -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $VbssRecords = Get-VideoAppSharingStreams -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $VideoRecords = Get-VideoRecords -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $AppShareRecords = Get-AppShareRecords -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $RMCRecords = Get-RMC -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $IMFEDRecords = Get-IMFederatedDomains -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]
    $UserRegistration = Get-LastUserRegistration -sessions $sessions -sipAddress $userHash.SipAddress.Split(":")[1]

    # Write out reports
    if ($AudioRecords) {
        $AudioRecords | Export-Csv -Path $Global:AudioReports -NoTypeInformation -Append
    }

    if ($ReliabilityRecords) {
        $ReliabilityRecords | Export-Csv -Path $Global:ReliabilityReports -NoTypeInformation -Append
    }

    if ($VbssRecords) {
        $VbssRecords | Export-Csv -Path $Global:VbssReports -NoTypeInformation -Append
    }

    if ($VideoRecords) {
        $VideoRecords | Export-Csv -Path $Global:VideoReports -NoTypeInformation -Append
    }

    if ($AppShareRecords) {
        $AppShareRecords | Export-Csv -Path $Global:AppShareReports -NoTypeInformation -Append
    }

    if ($RMCRecords) {
        $RMCRecords | Export-Csv -Path $Global:RMCReports -NoTypeInformation -Append
    }

    if ($IMFEDRecords) {
        $IMFEDRecords | Export-Csv -Path $Global:IMFEDReports -NoTypeInformation -Append
    }

    if ($UserRegistration) {
        $UserRegistration | Export-Csv -Path $Global:UserRegReports -NoTypeInformation -Append
    }
    
}

function Get-AudioRecords {
    [cmdletbinding()]
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -match "Audio" -and $_.QoeReport -ne $null}

    # Format Base Address IP Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        #process feedback tokens from array into string to present back as an object we can report on
        if ($_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0) {
            [array]$arrTokens = $_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0 | Select-Object Id #declare an array so we can get the count. if the count is only one, the trim statement preceding won't make sense so we need to handle this.
            if ($arrTokens.Count -gt 1) {
                $arrTokens = $arrTokens.id.trim() -join "," #output would show System.Object[] in the report otherwise so we need to convert them to a string.
            }
        }
        else {
            $arrTokens = ""
        }

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$Events += [PSCustomObject][ordered]@{
            SipAddress                                  = $sipAddress
            StartTime                                   = $_.StartTime
            EndTime                                     = $_.EndTime
            DialogId                                    = $_.DialogId
            Conference                                  = $_.ConferenceUrl
            CallerUri                                   = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri                                   = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerIP                                    = $ipaddr.Matches($Ip[0]).value
            CalleeIP                                    = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                                = $subnet.Matches($Ip[0]).value
            CalleeSubnet                                = $subnet.Matches($Ip[1]).value
            CallerUserAgent                             = $_.FromClientVersion
            CalleeUserAgent                             = $_.ToClientVersion
            MediaType                                   = $_.MediaTypesDescription
            MediaStartTime                              = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                                = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds                      = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds  ### Need to account for NULL
            AudioTransport                              = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).Transport
            CallerCaptureDevice                         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDev
            CallerCaptureDriver                         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDevDriver
            CallerRenderDevice                          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromRenderDev
            CallerRenderDriver                          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromRenderDevDriver
            CalleeCaptureDevice                         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDev
            CalleeCaptureDriver                         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDriver
            CalleeRenderDevice                          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToRenderDev
            CalleeRenderDriver                          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToRenderDevDriver
            CallerConnectivityIce                       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromConnectivityIce
            CalleeConnectivityIce                       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToConnectivityIce
            CallerVPN                                   = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromVPN
            CalleeVPN                                   = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToVPN
            CallerNetworkConnectionDetial               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromNetworkConnectionDetail
            CalleeNetworkConnectionDetail               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToNetworkConnectionDetail
            CallerReflexiveLocalIPAddr                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromReflexiveLocalIPAddr
            CalleeReflexiveLocalIPAddr                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToReflexiveLocalIPAddr
            CallerFromBssid                             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromBssid
            CalleeToBssid                               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToBssid
            CallerFromWifiDriverDeviceDesc              = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiDriverDeviceDesc
            CalleeToWifiDriverDeviceDesc                = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiDriverDeviceDesc
            CallerFromWifiDriverDeviceVersion           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiDriverDeviceVersion
            CalleeToWifiDriverDeviceVersion             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiDriverDeviceVersion
            CallerFromWifiRSSI                          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiRSSI
            CalleeToWifiRSSI                            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiRSSI
            CallerFromSSID                              = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromSSID
            CalleeToSSID                                = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToSSID
            CallerFromWifiChannel                       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiChannel
            CalleeToWifiChannel                         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiChannel
            CallerFromActivePowerProfile                = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromActivePowerProfile
            CalleeToActivePowerProfile                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToActivePowerProfile
            CallerFromWifiHandovers                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiHandovers
            CalleeToWifiHandovers                       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiHandovers
            CallerFromWifiChannelSwitches               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiChannelSwitches
            CalleeToWifiChannelSwitches                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiChannelSwitches
            CallerFromWifiChannelReassociations         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiChannelReassociations
            CalleeToWifiChannelReassociations           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiChannelReassociations
            CallerFromWifiRadioFrequency                = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiRadioFrequency
            CalleeToWifiRadioFrequency                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiRadioFrequency
            CallerFromWifiSignalStrength                = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromWifiSignalStrength
            CalleeToWifiSignalStrength                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToWifiSignalStrength
            CallerJitter                                = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrival -strDirection 'FROM-to-TO'
            CallerJitterMax                             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrivalMax -strDirection 'FROM-to-TO'
            CallerPacketLossRate                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRate -strDirection 'FROM-to-TO'
            CallerPacketLossRateMax                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRateMax -strDirection 'FROM-to-TO'
            CallerRoundTrip                             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTrip -strDirection 'FROM-to-TO'
            CallerRoundTripMax                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTripMax -strDirection 'FROM-to-TO'
            CallerRatioConcealedSamplesAvg              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RatioConcealedSamplesAvg -strDirection 'FROM-to-TO'
            CallerDegradationAvg                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam DegradationAvg -strDirection 'FROM-to-TO'
            CallerDegradationMax                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam DegradationMax -strDirection 'FROM-to-TO'
            CallerConcealedRatioMax                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam ConcealedRatioMax -strDirection 'FROM-to-TO'
            CallerAvgNetworkMOS                         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam OverallAvgNetworkMOS -strDirection 'FROM-to-TO'
            CallerSendListenMOS                         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam SendListenMOS -strDirection 'FROM-to-TO'
            CallerBandwidthEst                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam BandwidthEst -strDirection 'FROM-to-TO'
            CallerAudioFECUsed                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam AudioFECUsed -strDirection 'FROM-to-TO'
            CallerPayloadDescription                    = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PayloadDescription -strDirection 'FROM-to-TO'
            CallerStreamDirection                       = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam StreamDirection -strDirection 'FROM-to-TO'
            CalleeJitter                                = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrival -strDirection 'TO-to-FROM'
            CalleeJitterMax                             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrivalMax -strDirection 'TO-to-FROM'
            CalleePacketLossRate                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRate -strDirection 'TO-to-FROM'
            CalleePacketLossRateMax                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRateMax -strDirection 'TO-to-FROM'
            CalleeRoundTrip                             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTrip -strDirection 'TO-to-FROM'
            CalleeRoundTripMax                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTripMax -strDirection 'TO-to-FROM'
            CalleeRatioConcealedSamplesAvg              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RatioConcealedSamplesAvg -strDirection 'TO-to-FROM'
            CalleeDegradationAvg                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam DegradationAvg -strDirection 'TO-to-FROM'
            CalleeDegradationMax                        = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam DegradationMax -strDirection 'TO-to-FROM'
            CalleeConcealedRatioMax                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam ConcealedRatioMax -strDirection 'TO-to-FROM'
            CalleeAvgNetworkMOS                         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam OverallAvgNetworkMOS -strDirection 'TO-to-FROM'
            CalleeSendListenMOS                         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam SendListenMOS -strDirection 'TO-to-FROM'
            CalleeBandwidthEst                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam BandwidthEst -strDirection 'TO-to-FROM'
            CalleeAudioFECUsed                          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam AudioFECUsed -strDirection 'TO-to-FROM'
            CalleePayloadDescription                    = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PayloadDescription -strDirection 'TO-to-FROM'
            CalleeStreamDirection                       = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam StreamDirection -strDirection 'TO-to-FROM'
            CallerNetworkSendQualityEventRatio          = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkSendQualityEventRatio
            CallerNetworkReceiveQualityEventRatio       = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkReceiveQualityEventRatio
            CallerNetworkDelayEventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkDelayEventRatio
            CallerNetworkBandwidthLowEventRatio         = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkBandwidthLowEventRatio
            CallerCPUInsufficientEventRatio             = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).CPUInsufficientEventRatio
            CallerDeviceRenderNotFunctioningEventRatio  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceRenderNotFunctioningEventRatio
            CallerDeviceCaptureNotFunctioningEventRatio = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceCaptureNotFunctioningEventRatio
            CallerDeviceGlitchesEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceGlitchesEventRatio
            CallerDeviceLowSNREventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceLowSNREventRatio
            CallerDeviceLowSpeechLevelEventRatio        = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceLowSpeechLevelEventRatio
            CallerDeviceClippingEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceClippingEventRatio
            CallerDeviceEchoEventRatio                  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceEchoEventRatio
            CallerDeviceNearEndToEchoRatioEventRatio    = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceNearEndToEchoRatioEventRatio
            CallerDeviiceRenderZeroVolumeEventRatio     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviiceRenderZeroVolumeEventRatio
            CallerDeviceRenderMuteEventRatio            = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceRenderMuteEventRatio
            CallerDeviceMultipleEndpointsEventCount     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceMultipleEndpointsEventCount
            CallerDeviceHowlingEventCount               = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceHowlingEventCount
            CalleeNetworkSendQualityEventRatio          = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkSendQualityEventRatio
            CalleeNetworkReceiveQualityEventRatio       = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkReceiveQualityEventRatio
            CalleeNetworkDelayEventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkDelayEventRatio
            CalleeNetworkBandwidthLowEventRatio         = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkBandwidthLowEventRatio
            CalleeCPUInsufficientEventRatio             = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).CPUInsufficientEventRatio
            CalleeDeviceRenderNotFunctioningEventRatio  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceRenderNotFunctioningEventRatio
            CalleeDeviceCaptureNotFunctioningEventRatio = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceCaptureNotFunctioningEventRatio
            CalleeDeviceGlitchesEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceGlitchesEventRatio
            CalleeDeviceLowSNREventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceLowSNREventRatio
            CalleeDeviceLowSpeechLevelEventRatio        = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceLowSpeechLevelEventRatio
            CalleeDeviceClippingEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceClippingEventRatio
            CalleeDeviceEchoEventRatio                  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceEchoEventRatio
            CalleeDeviceNearEndToEchoRatioEventRatio    = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceNearEndToEchoRatioEventRatio
            CalleeDeviiceRenderZeroVolumeEventRatio     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviiceRenderZeroVolumeEventRatio
            CalleeDeviceRenderMuteEventRatio            = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceRenderMuteEventRatio
            CalleeDeviceMultipleEndpointsEventCount     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceMultipleEndpointsEventCount
            CalleeDeviceHowlingEventCount               = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceHowlingEventCount
            CallerSendSignalLevel                       = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $True}).SendSignalLevel
            CallerRecvSignalLevel                       = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $True}).RecvSignalLevel
            CallerAudioSpeakerGlitchRate                = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $True}).AudioSpeakerGlitchRate
            CallerAudioMicGlitchRate                    = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $True}).AudioMicGlitchRate
            CalleeSendSignalLevel                       = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $False}).SendSignalLevel
            CalleeRecvSignalLevel                       = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $False}).RecvSignalLevel
            CalleeAudioSpeakerGlitchRate                = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $False}).AudioSpeakerGlitchRate
            CalleeAudioMicGlitchRate                    = $_.QoeReport.AudioSignals.where( {$_.SubmittedByFromUser -eq $False}).AudioMicGlitchRate
        } 
        #}
    }
    return $Events

}

function Get-VideoRecords {
    [cmdletbinding()]
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -eq "[Audio][Video]" -or $_.MediaTypesDescription -eq "[Conference][Audio][Video]" -and $_.QoeReport -ne $null}

    # Format Base Address IP Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        #process feedback tokens from array into string to present back as an object we can report on
        if ($_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0) {
            [array]$arrTokens = $_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0 | Select-Object Id #declare an array so we can get the count. if the count is only one, the trim statement preceding won't make sense so we need to handle this.
            if ($arrTokens.Count -gt 1) {
                $arrTokens = $arrTokens.id.trim() -join "," #output would show System.Object[] in the report otherwise so we need to convert them to a string.
            }
        }
        else {
            $arrTokens = ""
        }

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$Events += [PSCustomObject][ordered]@{
            SipAddress                     = $sipAddress
            StartTime                      = $_.StartTime
            EndTime                        = $_.EndTime
            DialogId                       = $_.DialogId
            Conference                     = $_.ConferenceUrl
            CallerUri                      = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri                      = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerIP                       = $ipaddr.Matches($Ip[0]).value
            CalleeIP                       = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                   = $subnet.Matches($Ip[0]).value
            CalleeSubnet                   = $subnet.Matches($Ip[1]).value
            CallerUserAgent                = $_.FromClientVersion
            CalleeUserAgent                = $_.ToClientVersion
            MediaType                      = $_.MediaTypesDescription
            MediaStartTime                 = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                   = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds         = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds
            VideoTransport                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).Transport
            CallerCaptureDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromCaptureDev
            CallerCaptureDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromCaptureDevDriver
            CallerRenderDevice             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromRenderDev
            CallerRenderDriver             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromRenderDevDriver
            CalleeCaptureDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToCaptureDev
            CalleeCaptureDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToCaptureDriver
            CalleeRenderDevice             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToRenderDev
            CalleeRenderDriver             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToRenderDevDriver
            CallerConnectivityIce          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromConnectivityIce
            CalleeConnectivityIce          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToConnectivityIce
            CallerVPN                      = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromVPN
            CalleeVPN                      = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToVPN
            CallerNetworkConnectionDetial  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromNetworkConnectionDetail
            CalleeNetworkConnectionDetail  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToNetworkConnectionDetail
            CallerReflexiveLocalIPAddr     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromReflexiveLocalIPAddr
            CalleeReflexiveLocalIPAddr     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToReflexiveLocalIPAddr
            CallerJitter                   = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam JitterInterArrival -strDirection 'FROM-to-TO'
            CallerJitterMax                = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam JitterInterArrivalMax -strDirection 'FROM-to-TO'
            CallerPacketLossRate           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PacketLossRate -strDirection 'FROM-to-TO'
            CallerPacketLossRateMax        = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PacketLossRateMax -strDirection 'FROM-to-TO'
            CallerRoundTrip                = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RoundTrip -strDirection 'FROM-to-TO'
            CallerRoundTripMax             = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RoundTripMax -strDirection 'FROM-to-TO'
            CallerBandwidthEst             = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam BandwidthEst -strDirection 'FROM-to-TO'
            CallerPayloadDescription       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PayloadDescription -strDirection 'FROM-to-TO'
            CallerSendCodecTypes           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendCodecTypes -strDirection 'FROM-to-TO'
            CallerSendResolutionWidth      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendResolutionWidth -strDirection 'FROM-to-TO'
            CallerSendResolutionHeight     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendResolutionHeight -strDirection 'FROM-to-TO'
            CallerSendFrameRateAverage     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendFrameRateAverage -strDirection 'FROM-to-TO'
            CallerSendBitRateMaximum       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendBitRateMaximum -strDirection 'FROM-to-TO'
            CallerSendBitRateAverage       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendBitRateAverage -strDirection 'FROM-to-TO'    
            CallerSendVideoStreamsMax      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendVideoStreamsMax -strDirection 'FROM-to-TO'
            CallerRecvCodecTypes           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvCodecTypes -strDirection 'FROM-to-TO'
            CallerRecvResolutionWidth      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvResolutionWidth -strDirection 'FROM-to-TO'
            CallerRecvResolutionHeight     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvResolutionHeight -strDirection 'FROM-to-TO'
            CallerRecvFrameRateAverage     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvFrameRateAverage -strDirection 'FROM-to-TO'
            CallerRecvBitRateMaximum       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvBitRateMaximum -strDirection 'FROM-to-TO'
            CallerRecvBitRateAverage       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvBitRateAverage -strDirection 'FROM-to-TO'
            CallerCIFQualityRatio          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam CIFQualityRatio -strDirection 'FROM-to-TO'
            CallerVGAQualityRatio          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam VGAQualityRatio -strDirection 'FROM-to-TO'
            CallerHD720QualityRatio        = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam HD720QualityRatio -strDirection 'FROM-to-TO'
            CallerVideoPostFECPLR          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam VideoPostFECPLR -strDirection 'FROM-to-TO'
            CallerLowFrameRateCallPercent  = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowFrameRateCallPercent -strDirection 'FROM-to-TO'
            CallerLowBitRateCallPercent    = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowBitRateCallPercent -strDirection 'FROM-to-TO'
            CallerLowResolutionCallPercent = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowResolutionCallPercent -strDirection 'FROM-to-TO'
            CallerDynamicCapabilityPercent = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam DynamicCapabilityPercent -strDirection 'FROM-to-TO'
            CallerStreamDirection          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam StreamDirection -strDirection 'FROM-to-TO'
            CalleeJitter                   = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam JitterInterArrival -strDirection 'TO-to-FROM'
            CalleeJitterMax                = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam JitterInterArrivalMax -strDirection 'TO-to-FROM'
            CalleePacketLossRate           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PacketLossRate -strDirection 'TO-to-FROM'
            CalleePacketLossRateMax        = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PacketLossRateMax -strDirection 'TO-to-FROM'
            CalleeRoundTrip                = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RoundTrip -strDirection 'TO-to-FROM'
            CalleeRoundTripMax             = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RoundTripMax -strDirection 'TO-to-FROM'
            CalleeBandwidthEst             = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam BandwidthEst -strDirection 'TO-to-FROM'
            CalleePayloadDescription       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam PayloadDescription -strDirection 'TO-to-FROM'
            CalleeSendCodecTypes           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendCodecTypes -strDirection 'TO-to-FROM'
            CalleeSendResolutionWidth      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendResolutionWidth -strDirection 'TO-to-FROM'
            CalleeSendResolutionHeight     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendResolutionHeight -strDirection 'TO-to-FROM'
            CalleeSendFrameRateAverage     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendFrameRateAverage -strDirection 'TO-to-FROM'
            CalleeSendBitRateMaximum       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendBitRateMaximum -strDirection 'TO-to-FROM'
            CalleeSendBitRateAverage       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendBitRateAverage -strDirection 'TO-to-FROM'    
            CalleeSendVideoStreamsMax      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam SendVideoStreamsMax -strDirection 'TO-to-FROM'
            CalleeRecvCodecTypes           = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvCodecTypes -strDirection 'TO-to-FROM'
            CalleeRecvResolutionWidth      = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvResolutionWidth -strDirection 'TO-to-FROM'
            CalleeRecvResolutionHeight     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvResolutionHeight -strDirection 'TO-to-FROM'
            CalleeRecvFrameRateAverage     = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvFrameRateAverage -strDirection 'TO-to-FROM'
            CalleeRecvBitRateMaximum       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvBitRateMaximum -strDirection 'TO-to-FROM'
            CalleeRecvBitRateAverage       = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam RecvBitRateAverage -strDirection 'TO-to-FROM'
            CalleeCIFQualityRatio          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam CIFQualityRatio -strDirection 'TO-to-FROM'
            CalleeVGAQualityRatio          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam VGAQualityRatio -strDirection 'TO-to-FROM'
            CalleeHD720QualityRatio        = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam HD720QualityRatio -strDirection 'TO-to-FROM'
            CalleeVideoPostFECPLR          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam VideoPostFECPLR -strDirection 'TO-to-FROM'
            CalleeLowFrameRateCallPercent  = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowFrameRateCallPercent -strDirection 'TO-to-FROM'
            CalleeLowBitRateCallPercent    = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowBitRateCallPercent -strDirection 'TO-to-FROM'
            CalleeLowResolutionCallPercent = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam LowResolutionCallPercent -strDirection 'TO-to-FROM'
            CalleeDynamicCapabilityPercent = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam DynamicCapabilityPercent -strDirection 'TO-to-FROM'
            CalleeStreamDirection          = get-StreamParam -objParam $_.QoEReport.VideoStreams -strParam StreamDirection -strDirection 'TO-to-FROM'
        }
        #}
    }
    return $Events
}

function Get-AppShareRecords {
    [cmdletbinding()]
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -eq "[AppSharing]" -or $_.MediaTypesDescription -eq "[Conference][AppSharing]" -and $_.QoeReport -ne $null}

    # Format Base Address IP Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$Events += [PSCustomObject][ordered]@{
            SipAddress                          = $sipAddress
            StartTime                           = $_.StartTime
            EndTime                             = $_.EndTime
            DialogId                            = $_.DialogId
            Conference                          = $_.ConferenceUrl
            CallerUri                           = $_.FromUri
            CalleeUri                           = $_.ToUri
            CallerIP                            = $ipaddr.Matches($Ip[0]).value
            CalleeIP                            = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                        = $subnet.Matches($Ip[0]).value
            CalleeSubnet                        = $subnet.Matches($Ip[1]).value
            CallerUserAgent                     = $_.FromClientVersion
            CalleeUserAgent                     = $_.ToClientVersion
            MediaType                           = $_.MediaTypesDescription
            MediaStartTime                      = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                        = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds              = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds
            AppShareTransport                   = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).Transport
            CallerCaptureDevice                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromCaptureDev
            CallerCaptureDriver                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromCaptureDevDriver
            CallerRenderDevice                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromRenderDev
            CallerRenderDriver                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromRenderDevDriver
            CalleeCaptureDevice                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToCaptureDev
            CalleeCaptureDriver                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToCaptureDriver
            CalleeRenderDevice                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToRenderDev
            CalleeRenderDriver                  = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToRenderDevDriver
            CallerConnectivityIce               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromConnectivityIce
            CalleeConnectivityIce               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToConnectivityIce
            CallerVPN                           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromVPN
            CalleeVPN                           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToVPN
            CallerNetworkConnectionDetial       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromNetworkConnectionDetail
            CalleeNetworkConnectionDetail       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToNetworkConnectionDetail
            CallerReflexiveLocalIPAddr          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).FromReflexiveLocalIPAddr
            CalleeReflexiveLocalIPAddr          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "data"}).ToReflexiveLocalIPAddr
            CallerJitter                        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam JitterInterArrival -strDirection 'FROM-to-TO'
            CallerJitterMax                     = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam JitterInterArrivalMax -strDirection 'FROM-to-TO'
            CallerRoundTrip                     = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RoundTrip -strDirection 'FROM-to-TO'
            CallerRoundTripMax                  = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RoundTripMax -strDirection 'FROM-to-TO'
            CallerPacketUtilization             = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam PacketUtilization -strDirection 'FROM-to-TO'
            CallerAverageRectangleHeight        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam AverageRectangleHeight -strDirection 'FROM-to-TO'
            CallerAverageRectangleWidth         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam AverageRectangleWidth -strDirection 'FROM-to-TO'
            CallerRDPTileProcessingLatencyTotal = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RDPTileProcessingLatencyTotal -strDirection 'FROM-to-TO'
            CallerCaptureTileRateTotal          = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam CapturetileRateTotal -strDirection 'FROM-to-TO'
            CallerSpoiledTilePercentTotal       = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam SpoiledTilePercentTotal -strDirection 'FROM-to-TO'
            CallerScrapingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam ScrapingFrameRateTotal -strDirection 'FROM-to-TO'
            CallerIncomingTileRateTotal         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam IncomingTileRateTotal -strDirection 'FROM-to-TO'
            CallerIncomingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam IncomingFrameRateTotal -strDirection 'FROM-to-TO'
            CallerOutgoingTileRateTotal         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam OutgoingTileRateTotal -strDirection 'FROM-to-TO'
            CallerOutgoingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam OutgoingFrameRateTotal -strDirection 'FROM-to-TO'
            CallerStreamDirection               = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam StreamDirection -strDirection 'FROM-to-TO'
            CalleeJitter                        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam JitterInterArrival -strDirection 'TO-to-FROM'
            CalleeJitterMax                     = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam JitterInterArrivalMax -strDirection 'TO-to-FROM'
            CalleeRoundTrip                     = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RoundTrip -strDirection 'TO-to-FROM'
            CalleeRoundTripMax                  = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RoundTripMax -strDirection 'TO-to-FROM'
            CalleePacketUtilization             = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam PacketUtilization -strDirection 'TO-to-FROM'
            CalleeAverageRectangleHeight        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam AverageRectangleHeight -strDirection 'TO-to-FROM'
            CalleeAverageRectangleWidth         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam AverageRectangleWidth -strDirection 'TO-to-FROM'
            CalleeRDPTileProcessingLatencyTotal = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam RDPTileProcessingLatencyTotal -strDirection 'TO-to-FROM'
            CalleeCaptureTileRateTotal          = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam CapturetileRateTotal -strDirection 'TO-to-FROM'
            CalleeSpoiledTilePercentTotal       = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam SpoiledTilePercentTotal -strDirection 'TO-to-FROM'
            CalleeScrapingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam ScrapingFrameRateTotal -strDirection 'TO-to-FROM'
            CalleeIncomingTileRateTotal         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam IncomingTileRateTotal -strDirection 'TO-to-FROM'
            CalleeIncomingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam IncomingFrameRateTotal -strDirection 'TO-to-FROM'
            CalleeOutgoingTileRateTotal         = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam OutgoingTileRateTotal -strDirection 'TO-to-FROM'
            CalleeOutgoingFrameRateTotal        = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam OutgoingFrameRateTotal -strDirection 'TO-to-FROM'
            CalleeStreamDirection               = get-StreamParam -objParam $_.QoEReport.AppsharingStreams -strParam StreamDirection -strDirection 'TO-to-FROM'
            
        }
        #}
    }
    return $Events

}

function Get-AudioEvents {
    param (
        [object]$userHash,
        [datetime]$startTime,
        [datetime]$endTime
    )

    # Get User Sessions

    Write-Host "Processing data for $($userHash.SipAddress.Split(":")[1])" -ForegroundColor Red

    $sessions = getSession -SipAddress $userHash.SipAddress.Split(":")[1] -startTime $startTime -endTime $endTime

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -eq "[Conference][Audio]" -or $_.MediaTypesDescription -eq "[Audio]"} | Where-Object {$_.QoeReport -match "Session"}
    Write-Host "Sessions identified $($sessions.count)"

    # Format Base Address IP Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$Events += [PSCustomObject][ordered]@{
            StartTime                                   = $_.StartTime
            EndTime                                     = $_.EndTime
            DialogId                                    = $_.DialogId
            Conference                                  = $_.ConferenceUrl
            CallerUri                                   = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri                                   = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerIP                                    = $ipaddr.Matches($Ip[0]).value
            CalleeIP                                    = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                                = $subnet.Matches($Ip[0]).value
            CalleeSubnet                                = $subnet.Matches($Ip[1]).value
            CallerUserAgent                             = $_.FromClientVersion
            CalleeUserAgent                             = $_.ToClientVersion
            CallerLocalReflexive                        = $_.QoeReport.MediaLines.FromReflexiveLocalIPAddr
            CallerCaptureDevice                         = $_.QoeReport.MediaLines.FromCaptureDev
            CallerCaptureDeviceDriver                   = $_.QoeReport.MediaLines.FromCaptureDevDriver
            CallerRenderDevice                          = $_.QoeReport.MediaLines.FromRenderDev
            CallerRenderDeviceDriver                    = $_.QoeReport.MediaLines.FromRenderDevDriver
            CallerNetworkConnectionDetail               = $_.QoeReport.MediaLines.FromNetworkConnectionDetail
            CalleeLocalReflexive                        = $_.QoeReport.MediaLines.ToReflexiveLocalIPAddr
            CalleeCaptureDevice                         = $_.QoeReport.MediaLines.ToCaptureDev
            CalleeCaptureDeviceDriver                   = $_.QoeReport.MediaLines.ToCaptureDevDriver
            CalleeRenderDevice                          = $_.QoeReport.MediaLines.ToRenderDev
            CalleeRenderDeviceDriver                    = $_.QoeReport.MediaLines.ToRenderDevDriver
            CalleeNetworkConnectionDetail               = $_.QoeReport.MediaLines.ToNetworkConnectionDetail
            CallerNetworkSendQualityEventRatio          = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkSendQualityEventRatio
            CallerNetworkReceiveQualityEventRatio       = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkReceiveQualityEventRatio
            CallerNetworkDelayEventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkDelayEventRatio
            CallerNetworkBandwidthLowEventRatio         = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).NetworkBandwidthLowEventRatio
            CallerCPUInsufficientEventRatio             = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).CPUInsufficientEventRatio
            CallerDeviceRenderNotFunctioningEventRatio  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceRenderNotFunctioningEventRatio
            CallerDeviceCaptureNotFunctioningEventRatio = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceCaptureNotFunctioningEventRatio
            CallerDeviceGlitchesEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceGlitchesEventRatio
            CallerDeviceLowSNREventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceLowSNREventRatio
            CallerDeviceLowSpeechLevelEventRatio        = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceLowSpeechLevelEventRatio
            CallerDeviceClippingEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceClippingEventRatio
            CallerDeviceEchoEventRatio                  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceEchoEventRatio
            CallerDeviceNearEndToEchoRatioEventRatio    = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceNearEndToEchoRatioEventRatio
            CallerDeviiceRenderZeroVolumeEventRatio     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviiceRenderZeroVolumeEventRatio
            CallerDeviceRenderMuteEventRatio            = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceRenderMuteEventRatio
            CallerDeviceMultipleEndpointsEventCount     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceMultipleEndpointsEventCount
            CallerDeviceHowlingEventCount               = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $True}).DeviceHowlingEventCount
            CalleeNetworkSendQualityEventRatio          = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkSendQualityEventRatio
            CalleeNetworkReceiveQualityEventRatio       = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkReceiveQualityEventRatio
            CalleeNetworkDelayEventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkDelayEventRatio
            CalleeNetworkBandwidthLowEventRatio         = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).NetworkBandwidthLowEventRatio
            CalleeCPUInsufficientEventRatio             = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).CPUInsufficientEventRatio
            CalleeDeviceRenderNotFunctioningEventRatio  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceRenderNotFunctioningEventRatio
            CalleeDeviceCaptureNotFunctioningEventRatio = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceCaptureNotFunctioningEventRatio
            CalleeDeviceGlitchesEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceGlitchesEventRatio
            CalleeDeviceLowSNREventRatio                = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceLowSNREventRatio
            CalleeDeviceLowSpeechLevelEventRatio        = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceLowSpeechLevelEventRatio
            CalleeDeviceClippingEventRatio              = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceClippingEventRatio
            CalleeDeviceEchoEventRatio                  = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceEchoEventRatio
            CalleeDeviceNearEndToEchoRatioEventRatio    = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceNearEndToEchoRatioEventRatio
            CalleeDeviiceRenderZeroVolumeEventRatio     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviiceRenderZeroVolumeEventRatio
            CalleeDeviceRenderMuteEventRatio            = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceRenderMuteEventRatio
            CalleeDeviceMultipleEndpointsEventCount     = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceMultipleEndpointsEventCount
            CalleeDeviceHowlingEventCount               = $_.QoeReport.AudioClientEvents.where( {$_.SubmittedByFromUser -eq $False}).DeviceHowlingEventCount
        }
        #} 
    }
    return $Events
}

function Get-AudioQuality {
    param (
        [object]$userHash,
        [datetime]$startTime,
        [datetime]$endTime
    )

    # Get User Sessions
    Write-Host "Processing data for $($userHash.SipAddress.Split(":")[1])" -ForegroundColor Red
    $sessions = getSession -SipAddress $userHash.SipAddress.Split(":")[1] -startTime $startTime -endTime $endTime
     
    #$sessions = $sessions | Where-Object {$_.MediaTypesDescription -eq "[Conference][Audio]" -or $_.MediaTypesDescription -eq "[Audio]"} | Where-Object {$_.QoeReport -match 'Session'}
    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -match "Audio"} | Where-Object {$_.QoeReport -match 'Session'}
    Write-Host "Sessions identified $($sessions.count)"
    
    #Regex for Base Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$Events += [PSCustomObject][ordered]@{
            StartTime                        = $_.StartTime
            EndTime                          = $_.EndTime
            ConferenceUrl                    = $_.ConferenceUrl
            CallerUri                        = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri                        = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerUserAgent                  = $_.FromClientVersion
            CalleeUserAgent                  = $_.ToClientVersion
            CallerIP                         = $ipaddr.Matches($Ip[0]).value
            CalleeIP                         = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                     = $subnet.Matches($Ip[0]).value
            CalleeSubnet                     = $subnet.Matches($Ip[1]).value
            MediaStartTime                   = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                     = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds           = ($_.QoeReport.Session.MediaEndTime - $_.QoeReport.Session.MediaStartTime).Seconds
            AudioTransport                   = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).Transport
            AudioFromCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDev
            AudioFromCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDevDriver
            AudioToCaptureDevice             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDev
            AudioToCaptureDriver             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDriver
            AudioFromConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromConnectivityIce
            AudioToConnectivityIce           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToConnectivityIce
            AudioFromVPN                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromVPN
            AudioToVPN                       = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToVPN
            AudioFromNetworkConnectionDetial = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromNetworkConnectionDetail
            AudioToNetworkConnectionDetail   = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToNetworkConnectionDetail
            AudioFromReflexiveLocalIPAddr    = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromReflexiveLocalIPAddr
            AudioToReflexiveLocalIPAddr      = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToReflexiveLocalIPAddr
            CallerJitter                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrival -strDirection 'FROM-to-TO'
            CallerJitterMax                  = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrivalMax -strDirection 'FROM-to-TO'
            CallerPacketLossRate             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRate -strDirection 'FROM-to-TO'
            CallerPacketLossRateMax          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRateMax -strDirection 'FROM-to-TO'
            CallerRoundTrip                  = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTrip -strDirection 'FROM-to-TO'
            CallerRoundTripMax               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTripMax -strDirection 'FROM-to-TO'
            CallerAvgNetworkMOS              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam OverallAvgNetworkMOS -strDirection 'FROM-to-TO'
            CallerSendListenMOS              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam SendListenMOS -strDirection 'FROM-to-TO'
            CallerBandwidthEst               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam BandwidthEst -strDirection 'FROM-to-TO'
            CallerAudioFECUsed               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam AudioFECUsed -strDirection 'FROM-to-TO'
            CallerPayloadDescription         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PayloadDescription -strDirection 'FROM-to-TO'
            CalleeJitter                     = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrival -strDirection 'TO-to-FROM'
            CalleeJitterMax                  = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam JitterInterArrivalMax -strDirection 'TO-to-FROM'
            CalleePacketLossRate             = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRate -strDirection 'TO-to-FROM'
            CalleePacketLossRateMax          = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PacketLossRateMax -strDirection 'TO-to-FROM'
            CalleeRoundTrip                  = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTrip -strDirection 'TO-to-FROM'
            CalleeRoundTripMax               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam RoundTripMax -strDirection 'TO-to-FROM'
            CalleeAvgNetworkMOS              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam OverallAvgNetworkMOS -strDirection 'TO-to-FROM'
            CalleeSendListenMOS              = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam SendListenMOS -strDirection 'TO-to-FROM'
            CalleeBandwidthEst               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam BandwidthEst -strDirection 'TO-to-FROM'
            CalleeAudioFECUsed               = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam AudioFECUsed -strDirection 'TO-to-FROM'
            CalleePayloadDescription         = get-StreamParam -objParam $_.QoEReport.AudioStreams -strParam PayloadDescription -strDirection 'TO-to-FROM'
                
        }
        #}
    }

    return $Events
}

function Get-VideoAppSharingStreams {
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -eq "[Conference][Video][AppSharing]" -or $_.MediaTypesDescription -eq "[Video][AppSharing]"} | Where-Object {$_.QoeReport -match 'Session'}
        
    #Regex for Base Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
 
    $sessions| ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        # Get VBSS Stream Info
        $FallBackResults = RDPFallBack -ErrorReports $_.ErrorReports
        $FBStatus = $FallBackResults[0]
        $FBReason = $FallBackResults[1] 

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {


        [array]$Events += [PSCustomObject][ordered]@{
            SipAddress                    = $sipAddress
            StartTime                     = $_.StartTime
            EndTime                       = $_.EndTime
            DialogId                      = $_.DialogId
            Conference                    = $_.ConferenceUrl
            CallerUri                     = $_.FromUri
            CalleeUri                     = $_.ToUri
            CallerIP                      = $ipaddr.Matches($Ip[0]).value
            CalleeIP                      = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                  = $subnet.Matches($Ip[0]).value
            CalleeSubnet                  = $subnet.Matches($Ip[1]).value
            CallerUserAgent               = $_.FromClientVersion
            CalleeUserAgent               = $_.ToClientVersion
            MediaType                     = $_.MediaTypesDescription
            MediaStartTime                = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                  = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds        = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds
            CallerLocalReflexive          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromReflexiveLocalIPAddr
            CalleeLocalReflexive          = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToReflexiveLocalIPAddr
            CallerRelayIPAddr             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromRelayIPAddr
            CalleeRelayIPAddr             = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToRelayIPAddr
            CallerNetworkConnectionDetail = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromNetworkConnectionDetail
            CalleeNetworkConnectionDetail = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToNetworkConnectionDetail
            CallerTransport               = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).Transport
            CallerCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromCaptureDev
            CalleeCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToCaptureDev
            CallerCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromCaptureDevDriver
            CalleeCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToCaptureDevDriver
            CallerRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromRenderDev
            CalleeRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToRenderDev
            CallerRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromRenderDevDriver
            CalleeRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToRenderDevDriver
            CallerConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromConnectivityIce
            CalleeConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToConnectivityIce
            CallerMediaIP                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).FromIPAddr
            CalleeMediaIP                 = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "applicationsharing-video"}).ToIPAddr
            Jitter                        = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).JitterInterArrival
            JitterMax                     = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).JitterInterArrivalMax
            PacketLossRate                = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).PacketLossRate
            PacketLossRateMax             = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).PacketLossRateMax
            RoundTrip                     = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RoundTrip
            RoundTripMax                  = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RoundTripMax
            BandwidthEst                  = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).BandwidthEst
            PayloadDescription            = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).PayloadDescription
            SendCodecTypes                = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendCodecTypes
            SendResolutionWidth           = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendResolutionWidth
            SendResolutionHeight          = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendResolutionHeight
            SendFrameRateAverage          = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendFrameRateAverage
            SendBitRateAverage            = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendBitRateAverage
            SendBitRateMaximum            = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendBitRateMaximum
            SendVideoStreamsMax           = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).SendVideoStreamsMax
            RecvCodecTypes                = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvCodecTypes
            RecvResolutionWidth           = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvResolutionWidth
            RecvResolutionHeight          = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvResolutionHeight
            RecvFrameRateAverage          = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvFrameRateAverage
            RecvBitRateAverage            = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvBitRateAverage
            RecvBitRateMaximum            = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).RecvBitRateMaximum
            CIFQualityRatio               = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).CIFQualityRatio
            VGAQualityRatio               = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).VGAQualityRatio
            HD720QualityRatio             = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).HD720QualityRatio
            VideoPostFECPLR               = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).VideoPostFECPLR
            LowFrameRateCallPercent       = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).LowFrameRateCallPercent
            LowBitRateCallPercent         = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).LowBitRateCallPercent
            LowResolutionCallPercent      = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).LowResolutionCallPercent
            DynamicCapabilityPercent      = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).DynamicCapabilityPercent
            StreamDirection               = $_.QoEReport.VideoStreams.where( {$_.VideoMediaLineLabelText -eq "applicationsharing-video"}).StreamDirection
            FallBackStatus                = $FBStatus
            FallBackReason                = $FBReason

        } 
        #}
    }
    return $Events
}

function Get-IMFederatedDomains {
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    #$sessions = getSession -SipAddress $userHash.SipAddress.Split(":")[1] -startTime $startTime -endTime $endTime

    $sessions = $sessions | Where-Object {$_.MediaTypesDescription -match "IM"} 

    $sessions| ForEach-Object {

        $fromDomain = $_.FromUri.Split("@")[1]
        #$toDomain = $_.ToUri.Split("@")[1]
        $myDomain = $userHash.SipAddress.Split("@")[1]

        if ($myDomain -eq $fromDomain) {
            $federatedDomain = $_.ToUri.Split("@")[1]
        }
        else {
            $federatedDomain = $_.FromUri.Split("@")[1]
        }

        [array]$IMFED += [PSCustomObject][ordered]@{
            StartTime       = $_.StartTime
            EndTime         = $_.EndTime
            FromUri         = $_.FromUri
            ToUri           = $_.ToUri
            FederatedDomain = $federatedDomain
            Count           = 1

        }

    }

    return $IMFed

}

function Get-SetupOrDrops {
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.ErrorReports.DiagnosticId -ge 21 -and $_.ErrorReports.DiagnosticId -le 39}
   
    #Regex for Base Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $diagHeader = ""
        $CallerIP = ""

        # Verify this a user reported diagnostic error
        if ($_.ErrorReports.DiagnosticHeader -match "UserType=`"Callee`"" -and $_.ErrorReports.RequestType -eq "BYE") {
            
            # Check if this is a setup failure reported by client
            if ($_.ErrorReports.RequestType -eq "BYE" -and ($_.ErrorReports.DiagnosticId -ge 21 -and $_.ErrorReports.DiagnosticId -le 29)) {
                $diagHeader = $_.ErrorReports.Where( {$_.RequestType -eq "BYE" -and $_.DiagnosticHeader -match "UserType=`"Callee`""})
                
                if ($diagHeader) {
                    $CallerIP = $diagHeader.DiagnosticHeader.Split(";") -match "LocalSite" | Select-Object -First 1
                    #$CallerRflx = "Unknown"
                }
                else {
                    #$CallerIP = "Unknown"
                    #$CallerRflx = "Unknown"
                }

            }
            else {
                # Check if this is a mid-call failure reported by client
                if ($_.ErrorReports.RequestType -eq "BYE" -and ($_.ErrorReports.DiagnosticId -ge 31 -and $_.ErrorReports.DiagnosticId -le 39)) {
                    $diagHeader = $_.ErrorReports.Where( {$_.RequestType -eq "BYE" -and $_.DiagnosticHeader -match "UserType=`"Callee`""})

                    if ($diagHeader) {
                        $CallerIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
                        $CallerRflx = $diagHeader.DiagnosticHeader.Split(";") -match "LocalAddress" | Select-Object -First 1
                    }
                    else {
                        #$CallerIP = "Unknown"
                        #$CallerRflx = "Unknown"
                    }
                }
            }
        }
        # This should be a catch for any server reported mid-call failures
        elseif ($_.ErrorReports.DiagnosticHeader -notmatch "UserType=`"Callee`"" -and $_.ErrorReports.RequestType -eq "BYE") {
            
            $diagHeader = $_.ErrorReports.Where( {$_.RequestType -eq "BYE"})
            
            if ($diagHeader) {
                #$CallerIP = "Unknown"
                $CallerRflx = $diagHeader.DiagnosticHeader.Split(";") -match "RemoteAddress" | Select-Object -First 1
            }
            else {
                #$CallerIP = "Unknown"
                #$CallerRflx = "Unknown"
            }
        }
        else {
            #$CallerIP = "Unknown"
            #$CallerRflx = "Unknown"
        }


        if ($CallerIP) {
            $xCallerIP = $ipaddr.Matches($CallerIP).value
        }
        else {
            $xCallerIP = ""
        }

        if ($CallerRflx) {
            $xCallerRflx = $ipaddr.Matches($CallerRflx).value
        }
        else {
            $xCallerRflx = ""
        }

        # Check to see who reported the Failure
        # Get only request types of BYE method
        $BYEREQ = $_.ErrorReports.Where( {$_.RequestType -eq "BYE"})

        # Determine if it is Client or Server reported
        # If it is server, there will be a source attribute in the BYE
        if ($BYEREQ.DiagnosticHeader -match "source") {
            $Source = "Server"
        }
        else {
            $Source = "User"
        }

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$DiagErrors += [PSCustomObject][ordered]@{
            SipAddress             = $sipAddress
            StartTime              = $_.StartTime
            EndTime                = $_.EndTime
            DialogId               = $_.DialogId
            ConferenceUrl          = $_.ConferenceUrl
            Source                 = $Source
            MediaType              = $_.MediaTypesDescription
            MediaStartTime         = $_.QoeReport.Session.MediaStartTime
            MediaEndTime           = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds
            CallerUri              = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri              = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerUserAgent        = $_.FromClientVersion
            CalleeUserAgent        = $_.ToClientVersion
            CallerIp               = $xCallerIP
            CallerSubnet           = $subnet.Matches($xCallerIP).value
            CallerRflxIp           = $xCallerRflx
            CallerRflxSubnet       = $subnet.Matches($CallerRflx).value
            DiagnosticId           = $_.ErrorReports.Where( {$_.RequestType -eq "BYE"}).DiagnosticId | Select-Object -First 1
            Reason                 = $_.ErrorReports.Where( {$_.RequestType -eq "BYE"}).DiagnosticHeader.Split(";") -match "reason" | ForEach-Object {$_.split('"')[1]} | Select-Object -First 1

        }
        #}
    }
    
    return $DiagErrors
}

function Get-RMC {
    [cmdletbinding()]
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | Where-Object {$_.QoeReport.FeedbackReports.Count -gt 0}
   
    # Format Base Address IP Address
    $ipaddr = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
    $subnet = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){2}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"

    $sessions | ForEach-Object {

        $Ip = get-BaseAddr -objError $_.ErrorReports -isConference $(if ($_.ConferenceUrl -eq "") {$false} else {$true})

        #process feedback tokens from array into string to present back as an object we can report on
        if ($_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0) {
            [array]$arrTokens = $_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne 0 | Select-Object Id #declare an array so we can get the count. if the count is only one, the trim statement preceding won't make sense so we need to handle this.
            if ($arrTokens.Count -gt 1) {
                $arrTokens = $arrTokens.id.trim() -join "," #output would show System.Object[] in the report otherwise so we need to convert them to a string.
            }
        }
        else {
            $arrTokens = ""
        }

        # Check streams only belong to user being searched
        # Set From Uri to handle Server submitted QoE
        $FromUri = getFromUser -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -IsReceived $_.QoEReport.Session.IsFromReceived
        
        #if ($FromUri -eq $sipAddress) {

        [array]$RMCFeedback += [PSCustomObject][ordered]@{
            SipAddress                         = $sipAddress
            StartTime                          = $_.StartTime
            EndTime                            = $_.EndTime
            DialogId                           = $_.DialogId
            Conference                         = $_.ConferenceUrl
            CallerUri                          = getUserUri -Type Caller -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CalleeUri                          = getUserUri -Type Callee -FromUri $_.FromUri -ToUri $_.ToUri -SubmittedByFromUser $_.QoeReport.AudioSignals[0].SubmittedByFromUser
            CallerIP                           = $ipaddr.Matches($Ip[0]).value
            CalleeIP                           = $ipaddr.Matches($Ip[1]).value
            CallerSubnet                       = $subnet.Matches($Ip[0]).value
            CalleeSubnet                       = $subnet.Matches($Ip[1]).value
            CallerUserAgent                    = $_.FromClientVersion
            CalleeUserAgent                    = $_.ToClientVersion
            MediaType                          = $_.MediaTypesDescription
            MediaStartTime                     = $_.QoeReport.Session.MediaStartTime
            MediaEndTime                       = $_.QoeReport.Session.MediaEndTime
            MediaDurationInSeconds             = (New-TimeSpan -Start $_.QoeReport.Session.MediaStartTime -End $_.QoeReport.Session.MediaEndTime).TotalSeconds
            AudioTransport                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).Transport
            AudioCallerCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDev
            AudioCallerCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromCaptureDevDriver
            AudioCallerRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromRenderDev
            AudioCallerRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromRenderDevDriver
            AudioCalleeCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDev
            AudioCalleeCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToCaptureDriver
            AudioCalleeRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToRenderDev
            AudioCalleeRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToRenderDevDriver
            AudioCallerConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromConnectivityIce
            AudioCalleeConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToConnectivityIce
            AudioCallerVPN                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromVPN
            AudioCalleeVPN                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToVPN
            AudioCallerNetworkConnectionDetial = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromNetworkConnectionDetail
            AudioCalleeNetworkConnectionDetail = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToNetworkConnectionDetail
            AudioCallerReflexiveLocalIPAddr    = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).FromReflexiveLocalIPAddr
            AudioCalleeReflexiveLocalIPAddr    = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-audio"}).ToReflexiveLocalIPAddr
            VideoTransport                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).Transport
            VideoCallerCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromCaptureDev
            VideoCallerCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromCaptureDevDriver
            VideoCallerRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromRenderDev
            VideoCallerRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromRenderDevDriver
            VideoCalleeCaptureDevice           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToCaptureDev
            VideoCalleeCaptureDriver           = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToCaptureDriver
            VideoCalleeRenderDevice            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToRenderDev
            VideoCalleeRenderDriver            = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToRenderDevDriver
            VideoCallerConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromConnectivityIce
            VideoCalleeConnectivityIce         = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToConnectivityIce
            VideoCallerVPN                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromVPN
            VideoCalleeVPN                     = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToVPN
            VideoCallerNetworkConnectionDetial = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromNetworkConnectionDetail
            VideoCalleeNetworkConnectionDetail = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToNetworkConnectionDetail
            VideoCallerReflexiveLocalIPAddr    = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).FromReflexiveLocalIPAddr
            VideoCalleeReflexiveLocalIPAddr    = $_.QoeReport.MediaLines.where( {$_.MediaLineLabelText -eq "main-video"}).ToReflexiveLocalIPAddr
            Rating                             = $_.QoeReport.FeedBackReports.Rating
            FeedbackText                       = $_.QoeReport.FeedBackReports.FeedbackText
            Tokens                             = $arrTokens.Id
           
        }
        #}

    }

    return $RMCFeedback

}

#This functions grabs the last user registration based on distinct FromUri and FromClientVersion Combination.
# Function written by Brian Swagger
function Get-LastUserRegistration {
    [cmdletbinding()]
    param (
        [object]$sessions,
        [string]$sipAddress
    )

    $sessions = $sessions | where {$_.MediaTypesDescription -eq "[RegisterEvent]" -and $_.FromClientVersion -ne ""} | Select FromUri,FromClientVersion,StartTime

    #$sessions = $sessions | Where-Object {$_.FromClientVersion -ne ""}
    $sessions | ForEach-Object {


            [array]$Events += [PSCustomObject][ordered]@{
                #SipAddress                                  = $sipAddress
                StartTime                                   = $_.StartTime
                CallerUri                                   = $_.FromUri
                FromClientVersion                           = $_.FromClientVersion
               
            } 
        #}
    }
    
    $FinalEvents = $Events | Group-Object FromUri,FromClientVersion | Foreach-Object {$_.Group | Sort-Object StartTime | Select-Object -Last 1}
    return $FinalEvents

}
function sessionmgmt {
    param(
        [string]$SipAddress,
        [int]$duration,
        [datetime]$endTime
    )

    <#  
        Thank you to Matthew Grames for the huge help with this function
        This function will continue to divide by 2 until it returns less than 1000 sessions 
        for a given user over the time period provided.
    #>

    [int]$seconds = $duration / 2
    $sessions = @()
 
    # Session First Half
    $sfh = Get-CsUserSession -User $SipAddress -startTime $endTime.AddSeconds( - $duration) -endTime $endTime.AddSeconds( - $seconds) -ErrorAction Stop -WarningAction SilentlyContinue
        
    if ($sfh.count -ge 1000) {
        sessionmgmt -SipAddress $SipAddress -duration $seconds -endTime $endTime.AddSeconds( - $seconds) 
    }
    else {
        $sessions += $sfh
    }

    # Session Second Half
    $ssh = Get-CsUserSession -User $SipAddress -startTime $endTime.AddSeconds( - $seconds) -endTime $endTime -ErrorAction Stop -WarningAction SilentlyContinue

    if ($ssh.count -ge 1000) {
        sessionmgmt -SipAddress $SipAddress -duration $seconds -endTime $endTime
    }
    else {
        $sessions += $ssh
    }

    return $sessions

}

function getSession () {
    param (
        [string]$SipAddress,
        [datetime]$startTime,
        [datetime]$endTime,
        [int]$instance=0
    )

    $dateRange = $endTime - $startTime
    $instance++
    
    try {
        $sessions = Get-CsUserSession -User $SipAddress -startTime $startTime -endTime $endTime -ErrorAction Stop -WarningAction SilentlyContinue
 
        if ($sessions.count -ge 1000) {
            Write-Verbose "More than 1000 sessions identified.  Breaking up time range to retreive all sessions"
            $sessions = sessionmgmt -SipAddress $SipAddress -duration $dateRange.TotalSeconds -endTime $endTime
        }
        
    }
    catch {
        Write-Host "Failed to retreive user $($SipAddress) session. Attempting to repair the session and will try again" -ForegroundColor Yellow
        LogWrite -LogMessage $SipAddress
        LogWrite -LogMessage $_  | Format-List * -Force

        if ($instance -le 2) {
            Start-Sleep -Seconds 2
            Get-MySession -ForceUpdate
            getSession -SipAddress $SipAddress -startTime $startTime -endTime $endTime -instance $instance
        } else {
            Write-Host "Failing to process user $($SipAddress), aborting attempt." -ForegroundColor Yellow
            break
        }
    }

    return $sessions
}

function get-BaseAddr {
    param(
        [object]$objError,
        [bool]$isConference
    )

    # Global Function Variables
    [string]$CallerIP = "NA"
    [string]$CalleeIP = "NA"
    [array]$IpAddr = @()

   
    if ($isConference -eq $true) {
        # This will get Conference Caller Base Address
        if ($objError.DiagnosticHeader -match "UserType=`"Callee`"" -and $objError.RequestType -eq "BYE") {
            $diagHeader = $objError.Where( {$_.DiagnosticHeader -match "UserType=`"Callee`"" -and $_.RequestType -eq "BYE"})
            if ($diagHeader) {
                $CallerIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
            }
        }
           
        # This will get the Conference Callee IP
        if ($objError.DiagnosticHeader -match "component" -and $objError.DiagnosticHeader -match "BaseAddress") {
            $diagHeader = $objError.Where( {$_.DiagnosticHeader -match "component" -and $_.DiagnosticHeader -match "BaseAddress"})
            if ($diagHeader) {
                $CalleeIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
            }
        }
           
    }
    else {
        
        # check if this is P2P VBSS Streams
        if ($objError.Diagnosticheader -match "application-sharing-video") {
            
            # Get Caller Base Address
            if ($objError.DiagnosticHeader -match "UserType=`"Caller`"") {
                $diagHeader = $objError.Where( {$_.DiagnosticHeader -match "UserType=`"Caller`""} )
                $CallerIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
            }
              
            # Get Callee Base Address
            if ($objError.DiagnosticHeader -match "UserType=`"Callee`"") {
                $diagHeader = $objError.Where( {$_.DiagnosticHeader -match "UserType=`"Callee`""} )
                $CalleeIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
            }
            
        }
        else {

            # This will get P2P Callee Base Address
            if ($objError.RequestType -eq "INVITE") {
                $diagHeader = $objError.Where( {$_.RequestType -eq "INVITE"})
                if ($diagHeader) {
                    $CalleeIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
                }
            }
                
            # This will get P2P Caller Base Address
            if ($objError.DiagnosticHeader -match "UserType=`"Callee`"" -and $objError.RequestType -eq "BYE") {
                $diagHeader = $objError.Where( {$_.DiagnosticHeader -match "UserType=`"Callee`"" -and $_.RequestType -eq "BYE"})
                if ($diagHeader) {
                    $CallerIP = $diagHeader.DiagnosticHeader.Split(";") -match "BaseAddress" | Select-Object -First 1
                }
            }

        }
         
    }

    $IpAddr += $CallerIP
    $IpAddr += $CalleeIP
   
    return $IpAddr
        
}

function get-StreamParam {
    param (
        [object]$objParam,
        [string]$strParam,
        [string]$strDirection
    )

    if ($strDirection -eq "FROM-to-TO") {

        if ($objParam.where( {$_.StreamDirection -eq "FROM-to-TO"}).Count -gt 1) {
            if ($objParam[0].$strParam) {
                $strValue = $objParam[0].$strParam
            }
            else {
                $strValue = $objParam[1].$strParam
            }
        }
        else {
            $strValue = $objParam.where( {$_.StreamDirection -eq "FROM-to-TO"}).$strParam
        }
    }
    else {
        $strValue = $objParam.where( {$_.StreamDirection -eq "TO-to-FROM"}).$strParam
    }

    return $strValue

}

# To handle the situation when the server reports the QoE Records
# Example: If there is a mid-call failure, where the client is unable to report the session then server would report.
#           However in this case the From Uri and To Uri are flipped in the QoE Results
# This change will allow for the originating Caller to be represented in the From Uri regardless of who reported the QoE results.
function getUserUri {
    param (
        [string]$Type,
        [string]$FromUri,
        [string]$ToUri,
        [bool]$SubmittedByFromUser
    )

    if ($SubmittedByFromUser -eq $False) {
        if ($Type -eq "Caller") {
            #$Uri = $ToUri - pending investigation
            $Uri = $Fromuri
        }
        else {
            #$Uri = $FromUri - pending investigation
            $Uri = $ToUri
        }
    } 
    
    if ($SubmittedByFromUser -eq $True) {
        if ($Type -eq "Caller") {
            $Uri = $Fromuri
        }
        else {
            $Uri = $ToUri
        }
    }
    

    return $Uri
}

# Function used in processing of duplicate results
function getFromUser {
    param (
        [string]$Type,
        [string]$FromUri,
        [string]$ToUri,
        [bool]$IsReceived
    )

    if ($IsReceived -eq $False) {
        if ($Type -eq "Caller") {
            $Uri = $ToUri
        }
        else {
            $Uri = $FromUri
        }
    } 
    
    if ($IsReceived -eq $True) {
        if ($Type -eq "Caller") {
            $Uri = $Fromuri
        }
        else {
            $Uri = $ToUri
        }
    }
    

    return $Uri
}

# Check if RDP Fallback Occured
function RDPFallBack {
    param (
        [object]$ErrorReports
    )

    switch ($ErrorReports.DiagnosticId) {
        "51032" {
            $result = $true
            $reason = "Session fall back to RDP due to recording"
            break;
        }
        "51033" {
            $result = $true
            $reason = "Session fall back to RDP due to program sharing or window sharing"
            break;
        }
        "51034" {
            $result = $true
            $reason = "Session fall back to RDP due to control is given or control is granted"
            break;
        }
        "52540" {
            $result = $true
            $reason = "Session fall back to RDP due to initialization timeout"
            break;
        }
        "52541" {
            $result = $true
            $reason = "Session fall back to RDP due to video loss"
            break;
        }
        "21026" {
            $result = $true
            $reason = "Session fall back to RDP because a client that does not support VBSS has joined"
            break;
        }
        Default {
            $result = $false
            $reason = "Session was VBSS"
        }
    }

    return $result, $reason
}
