function Get-MySession {
    param (
        [switch]$ForceUpdate
    )
    
    # Check if SFBO credentials have been created
    if (!(Get-StoredCredential -Target "SFBO SCRIPT CRED")) {
        Write-Verbose -Message "No existing credentials found, prompting user to enter new credentials"
        # Prompt user for new credentials
        Get-Credential -Message "Enter your O365 Credentials" | New-StoredCredential -Target "SFBO SCRIPT CRED" | Out-Null
        $myCreds = Get-StoredCredential -Target "SFBO SCRIPT CRED"
    }
    else {
        # Get Stored Credentials
        $myCreds = Get-StoredCredential -Target "SFBO SCRIPT CRED"
    }

    # Clean up any existing broken sessions
    Get-PSSession | Where-Object {$_.ComputerName -like "*.online.lync.com" -and $_.State -eq "Broken"} | Remove-PSSession

    if (!($Global:SessionTime)) {
        # This must be a a new session, verify no other sessions exists
        $ckSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.online.lync.com" -and $_.State -ne "Broken"}
        
        if (!($ckSession)) {
            # No Session Timer Found and no Active Existing Sessions found, Create new Sessions
            $MySession = New-MySession -username $myCreds.UserName -password $myCreds.Password

            # Set Session Time
            $Global:SessionTime = $MySession.Time
        }
    }
         
    if ($Global:SessionTime -lt (Get-Date).AddMinutes(-45) -or $ForceUpdate) {
        # Session is older than 45 minutes, clean up timer and refresh the session
        $Global:SessionTime = [string]::Empty

        # Remove old PSSession
        Get-PSSession | Where-Object {$_.ComputerName -like "*.online.lync.com"} | Remove-PSSession

        # Pause before starting a new session
        Write-Verbose -Message "Pause 3 seconds after removing PSSession"
        Start-Sleep -Seconds 3
        
        # Establish new session
        Write-Verbose -Message "Updating the existing PSSession"
        $MySession = New-MySession -username $myCreds.UserName -password $myCreds.Password

        # Set Session Time
        $Global:SessionTime = $MySession.Time

    }

   
}

function New-MySession {
    Param (
        [string]$username,
        [securestring]$password
    )

    # Store passed Credentials in PSCredential Object
    $mycredentials = New-Object System.Management.Automation.PSCredential ($username, $password)

    try {
        # Create new PSSession
        $session = New-CsOnlineSession -Credential $mycredentials
        #$session = New-CsOnlineSession -Credential $mycredentials -OverrideAdminDomain "domain.onmicrosoft.com"
    }
    catch {
        $ErrorMessage = $_.Exception
        Write-Host "Failed to create a new PSSession, will pause for 5 seconds and attempt again" -ForegroundColor Yellow
        Write-Host "Error : $($ErrorMessage)" -ForegroundColor Red
        Start-Sleep -Seconds 5
        Get-MySession -ForceUpdate
    }
   
    Import-Module (Import-PSSession $session -AllowClobber) -Global

    # Create Custom Object to record time when session was created
    $sessionObject = [PSCustomObject][Ordered]@{
        ComputerName = $session.ComputerName
        InstanceId   = $session.InstanceId
        Id           = $session.Id
        Time         = Get-Date
    }

    return $sessionObject
}



