#  Mailbox Export
#  Version 1.1.0
#
# Run bulk mailbox exports as a scheduled task.
# Copyright (C) 2016 Simon Lehmann
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#

# Script version
$ver = 1.1.0

# Destination share (format: \\MACHINENAME\Path\To\Share)
$serverShare = "\\SERVER\Export"

$scriptStart = (Get-Date)

# Get today's date
$dirDate = Get-Date -format yyyy-MM-dd_hh-mm-ss

echo $dirDate

# Add Exchange PowerShell module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


function DeleteRequests($requests)
{
    if ($requests.count -gt 0) 
    {

        echo “$($requests.Count) requests to remove”
        foreach ($request in $requests) {
            Remove-MailboxExportRequest -Identity:$request -Confirm:$False
        }
        echo "Complete..."
    }
    else 
    {
        echo "No requests to remove"
    }
}


# ----- MAILBOX EXPORT REQUEST CLEANUP ----- #
$existingRequests = Get-MailboxExportRequest -Status Completed

DeleteRequests($existingRequests)


# ----- MAILBOX EXPORT REQUEST CREATION ----- #

# Create dated folder
New-Item $serverShare\$dirDate -type directory

# Get all mailboxes
$mailboxes = Get-Mailbox

# Export all mailboxes to PST files with names based on mailbox aliases (to use a different mailbox property replace the phrase “Alias” with its name):
$mailboxes|%{$_|New-MailboxExportRequest -FilePath $serverShare\$dirDate\$($_.Alias).pst}



# ----- MAILBOX EXPORT REQUEST REMOVAL ----- #

$allCompleted = $false

#echo $result
$pass = 0
$completed = 0
while($pass -ne 360)
    {
        $pass++
        echo "Mailbox Removal Pass $pass"
        if ($allCompleted -eq $false) {
            $allCompleted = $true
            
            # Get all mailbox export requests
            $results = Get-MailboxExportRequest 
            foreach ($result in $results) {
                # Check if current mailbox export request has completed
                if ($result.Status -eq "Completed") {
                    # Remove current mailbox export request
                    Remove-MailboxExportRequest -Identity:$result -Confirm:$False
                    $completed++
                }

                else {
                    echo "Incomplete!"
                    $allCompleted = $false
                }
            }
            echo "$completed of $($mailboxes.count) requests completed and removed..."
        }
        else {
            echo "about to break"
            break
            
        }
        # wait 15 seconds
        Start-Sleep -s 10
    }
echo "Broken"

$processed = $mailboxes.Count

$scriptEnd = (Get-Date)
$runTime = New-Timespan -Start $scriptStart -End $scriptEnd
$elapsedTime = “{0}:{1}:{2}” -f $runTime.Hours,$runtime.Minutes,$runTime.Seconds
echo "Elapsed Time: $elapsedTime"
# ----- SEND LOG EMAIL ----- #

$PSEmailServer = "localhost"
Send-MailMessage -From "someone@domain.com" -To "someone@domain.com" -Subject "Mailbox Export Successful" -Body "Mailbox export job completed successfully.`n`nProcessed $processed mailboxes.`n`nElapsed Time: $elapsedTime`n`nMailbox export request removal passes: $pass`n`nExchange Mailbox Exporter Version $ver"
