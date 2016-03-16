### Exchange 2013 Mailbox Exporter by Simon Lehmann ###
# Script version
$ver = 1.0.2

# Init
$server = "HNS-P1"
$share = "Exports"

$ScriptStart = (Get-Date)

# Get today's date
$date = Get-Date -format yyyy-MM-dd_hh-mm-ss

echo $date

# Add Exchange PowerShell module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

# ----- MAILBOX EXPORT REQUEST CREATION ----- #

# Create dated folder
New-Item \\$server\$share\$date -type directory

# Save all mailboxes to a variable (in my case it’s AllMailboxes):
$AllMailboxes = Get-Mailbox

# Export all mailboxes to PST files with names based on mailbox aliases (to use a different mailbox property replace the phrase “Alias” with its name):
$AllMailboxes|%{$_|New-MailboxExportRequest -FilePath \\$server\$share\$date\$($_.Alias).pst}

# ----- MAILBOX EXPORT REQUEST REMOVAL ----- #

$allCompleted = $false

#echo $result
$i = 0
while($i -ne 60)
    {
        $i++
        echo "Mailbox Removal Pass $i"
        if ($allCompleted -eq $false) {
            $allCompleted = $true
            
            # Get all mailbox export requests
            $results = Get-MailboxExportRequest 
            foreach ($result in $results) {
                # Check if current mailbox export request has completed
                if ($result.Status -eq "Completed") {
                    # Remove current mailbox export request
                    Remove-MailboxExportRequest -Identity:$result -Confirm:$False
                }

                else {
                    echo "Incomplete!"
                    $allCompleted = $false
                }
            }
        }
        else {
            echo "about to break"
            break
            
        }
        # wait 15 seconds
        Start-Sleep -s 10
    }
echo "Broken"

$processed = $AllMailboxes.Count

$ScriptEnd = (Get-Date)
$RunTime = New-Timespan -Start $ScriptStart -End $ScriptEnd
$elapsedTime = “{0}:{1}:{2}” -f $RunTime.Hours,$Runtime.Minutes,$RunTime.Seconds
echo "Elapsed Time: $elapsedTime"
# ----- SEND LOG EMAIL ----- #

$PSEmailServer = "localhost"
Send-MailMessage -From "Syslog@hostnserve.com.au" -To "Syslog@hostnserve.com.au" -Subject "Mailbox Export Successful" -Body "Mailbox export job completed successfully.`n`nProcessed $Processed mailboxes.`n`nElapsed Time: $elapsedTime`n`nMailbox export request removal passes: $i`n`nExchange Mailbox Exporter Version $ver"