<#
                .PARAMETER  DateStart (MM/DD/YYYY)
                                Start date for search.

                .PARAMETER  DateEnd (MM/DD/YYYY)
                                End date for search.

                .PARAMETER  Subject
                                Subject of message to search.

                .PARAMETER  SenderEmailAddress
                                Email address of sender

                .PARAMETER  RecipientEmailAddress
                                Email address of recipient

                .PARAMETER  Message_ID
                                Message ID of the email

                .PARAMETER  FromIPAddress
                                IP address from which the email was sent
 
                .PARAMETER  OutputFile
                                Name of CSV file to populate with results.
#>

Write-Host "All parameters are optional. If you do not wish to use a certain parameter in your search, just press " -ForegroundColor Cyan -NoNewline
Write-Host "[Enter] " -ForegroundColor Green -NoNewline
Write-Host "to move on to the next one.`n" -ForegroundColor Cyan
Write-Host "This script outputs data to a CSV file. Please specify a filename when prompted`n" -ForegroundColor Red

Write-Host "Enter search start date (MM/DD/YYYY): " -ForegroundColor Yellow -NoNewline
$DateStart = Read-Host
Write-Host "Enter search end date (MM/DD/YYYY): " -ForegroundColor Yellow -NoNewline
$DateEnd = Read-Host
Write-Host "Enter the subject of the email: " -ForegroundColor Yellow -NoNewline
$Subject = Read-Host
Write-Host "Enter the sender email address: " -ForegroundColor Yellow -NoNewline
$SenderEmailAddress = Read-Host
Write-Host "Enter the recipient email address: " -ForegroundColor Yellow -NoNewline
$RecipientEmailAddress = Read-Host
Write-Host "Enter the Message-ID: " -ForegroundColor Yellow -NoNewline
$Message_ID = Read-Host
Write-Host "Enter the From IP Address: " -ForegroundColor Yellow -NoNewline
$FromIPAddress = Read-Host
Write-Host "Enter the output filename: " -ForegroundColor Yellow -NoNewline
$OutputFile = Read-Host

$htParams = @{
  
}

if($DateStart) 
{ 
    $htParams.StartDate = $DateStart 
}

if($DateEnd) 
{ 
$htParams.EndDate = $DateEnd 
}

if($SenderEmailAddress) 
{ 
$htParams.SenderAddress = $SenderEmailAddress 
}

if($RecipientEmailAddress) 
{ 
$htParams.RecipientAddress = $RecipientEmailAddress 
}

if($Message_ID) 
{ 
$htParams.MessageID = $Message_ID 
}

if($FromIPAddress) 
{ 
$htParams.FromIP = $FromIPAddress 
}
 
$FoundCount = 0
 
For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
{
    $Messages = Get-MessageTrace @htParams -PageSize 5000 -Page $i
 
    If($Messages.count -gt 0)
    {
        $Status = $Messages[-1].Received.ToString("MM/dd/yyyy HH:mm") + " - " + $Messages[0].Received.ToString("MM/dd/yyyy HH:mm") + "  [" + ("{0:N0}" -f ($i*5000)) + " Searched | " + $FoundCount + " Found]"
 
        Write-Progress -activity "Checking Messages (Up to 5 Million)..." -status $Status
 
        If(!$Subject)
        {
            $Entries = $Messages | Select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, Size, MessageId
            $Entries | Export-Csv $OutputFile -NoTypeInformation -Append
 
            $FoundCount += $Entries.Count
        }
        Else
        {
         $Entries = $Messages | Where {$_.Subject -like $Subject} | Select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, Size, MessageId
         $Entries | Export-Csv $OutputFile -NoTypeInformation -Append
 
         $FoundCount += $Entries.Count
        }
    }
    Else
    {
        Break
    }
} 
 
Write-Host $FoundCount "Entries Found & Logged In" $OutputFile
