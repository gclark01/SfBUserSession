<#
.SYNOPSIS
Process users in batches of 500

.DESCRIPTION
This function with take in an array of users and then group them into batches of 500.  This is returned as a Hash {Key: Integer, Value : Skype Users})

.NOTES
This function was borrowed from cxdCallData, available on PoweShell Gallery.
Credit goes to Jason Shave for putting this process together to handle processing of users in batches "buckets"
#>
function Get-BatchUsers {
    param (
        [array]$Users
    )

    #Function Variables
    $i = 0
    $_UserBatch = @{}

    # Get Total Users
    [int]$UCount = $Users.Count
    Write-Verbose -Message "Adding $($UCount) users into group(s)...."

    # Group Users
    $UserBatch = [System.Math]::Ceiling($UCount / 500)

    $Users | ForEach-Object {
        $_UserBatch[$i % $UserBatch] += @($_)
        $i++
    }

    Write-Verbose "Placed $($UCount) users into $($_UserBatch.Count) groups."

    return $_UserBatch
}

