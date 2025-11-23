# MFA-DomainSleuth - Microsoft Graph Edition
# This version uses Microsoft Graph API to work with new Outlook and Microsoft 365

# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
    Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Yellow
    Write-Host "Installing Microsoft.Graph.Mail module..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Mail -Scope CurrentUser -Force
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

# Import required modules
Import-Module Microsoft.Graph.Mail
Import-Module Microsoft.Graph.Authentication

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host "You will be prompted to sign in with your Microsoft 365 account." -ForegroundColor Yellow

try {
    # Connect with necessary permissions
    Connect-MgGraph -Scopes "Mail.Read", "Mail.ReadBasic" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Prompt the user for the number of days to check
do {
    $daysToCheck = Read-Host "Enter the number of days to check back on emails"
} while ($daysToCheck -notmatch '^\d+$' -or $daysToCheck -le 0)

do {
    $showUnsupported = Read-Host "Show domains that do not support MFA? (yes/no)"
} while ($showUnsupported -notmatch '^(yes|no)$')

# Prompt for folder selection (optional - default is Inbox)
Write-Host "`nAvailable folder options:" -ForegroundColor Cyan
Write-Host "1. Inbox (default)"
Write-Host "2. Sent Items"
Write-Host "3. All Mail"
Write-Host "4. Custom folder (enter folder name)"
$folderChoice = Read-Host "Select folder option (1-4, default: 1)"

$folderName = "Inbox"
switch ($folderChoice) {
    "2" { $folderName = "SentItems" }
    "3" { $folderName = "AllMail" }
    "4" { $folderName = Read-Host "Enter custom folder name" }
    default { $folderName = "Inbox" }
}

# Create a date object for the specified number of days ago
$startDate = (Get-Date).AddDays(-[int]$daysToCheck)
$filterDate = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Create hashtable for tracking processed domains
$processedDomains = @{}

# Get the directory of the current script
$scriptPath = Split-Path -Parent $PSCommandPath
if (-not $scriptPath) {
    $scriptPath = Get-Location
}

# Fetch JSON data of 2FA enabled/disabled websites
Write-Host "Fetching MFA/2FA directory data..." -ForegroundColor Cyan
try {
    $siteData = Invoke-RestMethod -Uri "https://api.2fa.directory/v3/all.json"
    Write-Host "Successfully loaded 2FA directory data." -ForegroundColor Green
} catch {
    Write-Host "Failed to fetch 2FA directory data: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Initialize HTML output with a modern CSS styling
$script:htmlOutput = @"
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 20px;
        }
        h1 {
            color: #333;
        }
        .info {
            background-color: #e7f3fe;
            border-left: 6px solid #2196F3;
            padding: 10px;
            margin-bottom: 20px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .mfaEnabled {
            background-color: #8bc34a;
        }
        .mfaDisabled {
            background-color: #ff4444;
            color: white;
        }
    </style>
</head>
<body>
    <h1>MFA-DomainSleuth Report</h1>
    <div class="info">
        <strong>Powered by Microsoft Graph API</strong><br>
        Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")<br>
        Date Range: Last $daysToCheck days<br>
        Folder: $folderName
    </div>
    <table>
        <tr>
            <th>Matched Domain</th>
            <th>Supports MFA?</th>
            <th>Email Address</th>
            <th>Documentation</th>
            <th>MFA Methods</th>
            <th>Custom Software</th>
            <th>Custom Hardware</th>
            <th>Recovery URL</th>
            <th>Additional Domains</th>
            <th>Keywords</th>
            <th>Notes</th>
        </tr>
"@

# Function to process each email
function ProcessEmail($email) {
    try {
        # Extract sender information - Graph API provides better structured data
        $senderEmail = $null

        if ($email.From.EmailAddress.Address) {
            $senderEmail = $email.From.EmailAddress.Address
        } elseif ($email.Sender.EmailAddress.Address) {
            $senderEmail = $email.Sender.EmailAddress.Address
        }

        if (-not $senderEmail) {
            return
        }

        # Extract the domain from the sender's email address
        $senderEmailParts = $senderEmail -split "@"
        if ($senderEmailParts.Count -lt 2) {
            return
        }

        $senderDomain = $senderEmailParts[-1].ToLower()

        # Skip processing if the domain matches the specified pattern
        if ($senderDomain -like "*EXCHANGELABS*" -or $senderDomain -like "*onmicrosoft.com*") {
            return
        }

        # Add domain to hashtable and directly write to HTML output if it's not already there
        if (!$processedDomains.ContainsKey($senderDomain)) {
            $processedDomains[$senderDomain] = $true
            $matchedDomain = $null

            # Determine if domain supports MFA and get documentation link
            $siteInfo = $siteData | Where-Object {
                if ($senderDomain -like "*.$($_[1].domain)" -or $senderDomain -eq $_[1].domain) {
                    $script:matchedDomain = $_[1].domain
                    return $true
                } elseif ($_[1].'additional-domains' | Where-Object { $senderDomain -like "*.$_" -or $senderDomain -eq $_ }) {
                    $script:matchedDomain = $_
                    return $true
                } else {
                    return $false
                }
            }

            if ($siteInfo) {
                $mfaSupport = "Yes"
                $mfaClass = "mfaEnabled"
                $SenderEmailAddress = $senderEmail
                $documentationLink = "<a href='$($siteInfo[1].documentation)' target='_blank'>Documentation</a>"
                $tfaMethods = $siteInfo[1].tfa -join ', '
                $customSoftware = $siteInfo[1].'custom-software' -join ', '
                $customHardware = $siteInfo[1].'custom-hardware' -join ', '
                $recoveryURL = $siteInfo[1].recovery
                $additionalDomains = $siteInfo[1].'additional-domains' -join ', '
                $keywords = $siteInfo[1].keywords -join ', '
                $notes = $siteInfo[1].notes
            } else {
                $mfaSupport = "No"
                $mfaClass = "mfaDisabled"
                $SenderEmailAddress = $senderEmail
                $documentationLink = ""
                $tfaMethods = ""
                $customSoftware = ""
                $customHardware = ""
                $recoveryURL = ""
                $additionalDomains = ""
                $keywords = ""
                $notes = ""
            }

            if (($mfaSupport -eq "Yes") -or ($showUnsupported -eq "yes" -and $mfaSupport -eq "No")) {
                $script:htmlOutput += @"
<tr>
    <td><a href="https://$matchedDomain" target="_blank">$matchedDomain</a></td>
    <td class='$mfaClass'>$mfaSupport</td>
    <td>$SenderEmailAddress</td>
    <td>$documentationLink</td>
    <td>$tfaMethods</td>
    <td>$customSoftware</td>
    <td>$customHardware</td>
    <td><a href="$recoveryURL" target="_blank">$recoveryURL</a></td>
    <td>$additionalDomains</td>
    <td>$keywords</td>
    <td>$notes</td>
</tr>
"@
            }
        }
    } catch {
        Write-Verbose "Error processing email: $_"
    }
}

# Retrieve emails using Microsoft Graph API
Write-Host "`nFetching emails from $folderName..." -ForegroundColor Cyan

try {
    # Build the filter query for date range
    $filter = "receivedDateTime ge $filterDate"

    # Get emails from the specified folder with pagination
    $emails = @()
    $pageSize = 100

    Write-Host "Retrieving emails (this may take a moment)..." -ForegroundColor Yellow

    # Get initial batch of messages
    $messages = Get-MgUserMessage -UserId "me" -Filter $filter -Top $pageSize -All

    $totalEmails = $messages.Count
    Write-Host "Found $totalEmails emails to process." -ForegroundColor Green

    # Process emails with progress bar
    $i = 0
    foreach ($email in $messages) {
        $i++
        Write-Progress -Activity "Processing Emails" -Status "$i out of $totalEmails Emails Processed" -PercentComplete (($i / $totalEmails) * 100)
        ProcessEmail $email
    }

    Write-Progress -Activity "Processing Emails" -Completed

} catch {
    Write-Host "Error retrieving emails: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Close HTML tags
$script:htmlOutput += @"
    </table>
    <div style="margin-top: 20px; color: #666; font-size: 12px;">
        <p>Total unique domains found: $($processedDomains.Count)</p>
        <p>Report generated using Microsoft Graph API</p>
    </div>
</body>
</html>
"@

# Write HTML output to file
$outputFile = Join-Path $scriptPath "MFA-DomainSleuth.html"
$script:htmlOutput | Out-File -FilePath $outputFile -Encoding UTF8

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Report completion
Write-Host "`nCompleted!" -ForegroundColor Green
Write-Host "Processed $totalEmails emails within the last $daysToCheck days." -ForegroundColor Cyan
Write-Host "Found $($processedDomains.Count) unique sender domains." -ForegroundColor Cyan
Write-Host "Report saved to: $outputFile" -ForegroundColor Cyan

# Optionally open the report
$openReport = Read-Host "`nWould you like to open the report now? (yes/no)"
if ($openReport -eq "yes") {
    Start-Process $outputFile
}
