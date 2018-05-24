# SfBUserSession
Based on Get-CsUserSession cmdlet to pull back QoE data from SfB O365

## Example Syntax
get-userstreams.ps1 -Type Report -startTime 04/01/2018 -endTime 04/15/2018 -ReportPath -c:\temp\reports

Optional Parameter
-CsvUsers c:\temp\specific-users.csv

## CSV Format
SipAddress<br>
sip:user1@contoso.com<br>
sip:user2@contoso.com<br>
sip:user3@contoso.com<br>

## Output
This will output six separate reports to the report path you provided when you executed the script
<ul>
    <li>Application Sharing</li>
    <li>Audio</li>
    <li>Reliability</li>
    <li>Rate My Call</li>
    <li>VBSS</li>
    <li>Video</li>
</ul>
