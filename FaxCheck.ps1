<#------------------------------------------------------------------------------
    Jason McClary
    mcclarj@mail.amc.edu
    29 Sep 2016
    30 Sep 2016 - Cleaned up email output

    
    Description:
    Check EDM_Fax folder and alert if new files are not being written
    
    Arguments:
    None
        
    Tasks:
    - Check create time of newest file
    - Alert if file is over 4 hours older

    References:
    http://stackoverflow.com/questions/9675658/powershell-get-childitem-most-recent-file-in-directory



--------------------------------------------------------------------------------
                                CONSTANTS
------------------------------------------------------------------------------#>
set-variable folderToCheck -option Constant -value "E:\ED_Fax"
set-variable monitorLog -option Constant -value "C:\Scripts\ED_Fax\FaxWarn.txt"
set-variable emailSendList -option Constant -value "C:\Scripts\ED_Fax\emailSendList.txt"

set-variable emailTSList -option Constant -value "<p>Some troubleshooting steps:<ul>
<li><b>Check if any faxes should have been received during this time frame</b></li>
<li>Check if fax has paper loaded</li>
<li>Check fax settings for DHCP, DNS and path to folder (\\ESERVER\ED_Fax)</li>
<li>Try to scan a copy to that folder</li>
<li>Check if share is accessible (\\ESERVER\ED_Fax)</li>
</ul></p>
"


<#------------------------------------------------------------------------------
                                Script Variables
------------------------------------------------------------------------------#>
$howOld = 2.5  # number of hours old a file should be to trigger alarm
$PSEmailServer = "SMTP_Server"
$sendFrom = "fromEmail"
$sendTo =  Get-Content $emailSendList
$mailPriority = "Normal"
$mailSubject = "NOT SET"
$mailBody = "EMPTY"

<#------------------------------------------------------------------------------
                                FUNCTIONS
------------------------------------------------------------------------------#>

    
<#------------------------------------------------------------------------------
                                    MAIN
------------------------------------------------------------------------------#>
# Check for monitor log file - if not there set up a blank
$fileName = ""
$alertCount = 0
IF (!(Test-Path $monitorLog)){
    $fileName > $monitorLog
    $alertCount >> $monitorLog
}

# Get the newest PDF in the folder
$file = get-childitem $folderToCheck *.pdf | sort LastWriteTime | select -last 1

# Calculate how many hours old the file is
$fileAge = (New-TimeSpan -Start ($file.LastWriteTime) -End $(Get-Date)).TotalHours

# Load log file for previous alerts
$current = Get-Content $monitorLog
$fileName = $current[0]
$alertCount = [int]$current[1]

IF ($fileAge -gt $howOld) {
    # The file is too old so do this...

    IF ($alertCount -gt 0){
        IF ($current[0] -eq $file.Name){
            $mailSubject = "ED Fax Folder Alert" # Alerts for same file

            switch ($alertCount){
                1       {$mailBody = "<p>This is the second e-mail about the ED fax machine not receiving a fax.</p>"}
                2       {$mailBody = "<p>This is the third e-mail about the ED fax machine not receiving a fax."}
                3       {$mailBody = "<p>This is the final e-mail about the ED fax machine.</p> <i>Further warnings will be suppressed. An all clear email will send once a new fax arrives.</i>"
                         $mailPriority = "High" }
                default {$mailBody = ""
                         $alertCount = 5 }
            }

            $file.name > $monitorLog
            $alertCount++
            $alertCount >> $monitorLog
            
        } ELSE { # Alert if a new fax arrives since an old one times out but it also times out. Should never be used if check interval is less then age of fax.
            $mailSubject = "ED Fax Folder ** New Warning **"
            $mailBody = "<p>A new fax has arrived since the last warning but now no further faxes have arrived.</p>"
            $file.name > $monitorLog
            $alertCount = 1
            $alertCount >> $monitorLog
        }

    } ELSE {
        $mailSubject = "ED Fax Folder Warning"
        $mailBody = "<p>This is the first e-mail about the ED fax machine not saving a fax.</p>"
        $file.name > $monitorLog
        $alertCount++
        $alertCount >> $monitorLog
    }
    $hoursSince =  [math]::Truncate($fileAge)
    $minutesSince = [math]::Truncate(($fileAge - $hoursSince) * 60)
    switch ($hoursSince) {
        1       { $hoursSince = "1 hour"}
        default { $hoursSince = "$hoursSince hours"}
    }
    switch ($minutesSince) {
        0       { $minutesSince = ""}
        1       { $minutesSince = " 1 minute"}
        default { $minutesSince = " $minutesSince minutes"}
    }

    $mailBody += "<p>It has been $hoursSince$minutesSince since the last fax was received.</p><p><b><font color=blue>$($file.name)</font> was the last fax recieved at <font color=red>$(Get-Date -Date $file.LastWriteTime -Format g)</font>.</b></p>"
    $mailBody += $emailTSList
    IF ($alertCount -lt 5) {Send-MailMessage -To $sendTo -Subject $mailSubject -Body $mailBody -BodyAsHtml -From $sendFrom -Priority $mailPriority}
} ELSE {
    # Fax is new enough so do this...
    IF ($alertCount -gt 0){
        $mailSubject = "ED Fax Folder All Clear"
        $mailBody = "The previous alert for <b><font color=blue>$fileName</font></b> has cleared as a new fax was received at <b><font color=red>$(Get-Date -Date $file.LastWriteTime -Format g)</font></b>."
        Send-MailMessage -To $sendTo -Subject $mailSubject -Body $mailBody -BodyAsHtml -From $sendFrom -Priority $mailPriority
    }
    # Set log back to empty/ 0
    $fileName = ""
    $alertCount = 0
    $fileName > $monitorLog
    $alertCount >> $monitorLog
}