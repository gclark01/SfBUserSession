function LogWrite {
    Param (
        [string]$LogMessage
    )

    Add-Content "$($Global:LogPath)\usersession.log" -value $LogMessage
}