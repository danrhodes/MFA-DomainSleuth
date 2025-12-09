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

# Prompt for user email address
Write-Host "Enter your email address (the account you want to scan for emails):" -ForegroundColor Cyan
$userEmail = Read-Host "Email address"

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
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
    $showUnsupported = Read-Host "Show domains that do not support MFA? (y/n)"
} while ($showUnsupported -notmatch '^(y|yes|n|no)$')
$showUnsupported = if ($showUnsupported -match '^y') { "yes" } else { "no" }

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
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>MFA-DomainSleuth Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            position: relative;
            overflow: hidden;
        }

        .header::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
        }

        .header h1 {
            font-size: 2.5em;
            font-weight: 700;
            margin-bottom: 10px;
            position: relative;
            z-index: 1;
        }

        .header p {
            font-size: 1.1em;
            opacity: 0.9;
            position: relative;
            z-index: 1;
        }

        .info-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px 40px;
            background: #f8f9fa;
        }

        .info-card {
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
            border-left: 4px solid #667eea;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .info-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.1);
        }

        .info-card strong {
            display: block;
            color: #667eea;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }

        .info-card span {
            font-size: 1.3em;
            font-weight: 600;
            color: #2d3748;
        }

        .table-container {
            padding: 40px;
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        thead {
            display: none;
        }

        tbody tr {
            display: block;
            margin-bottom: 20px;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            overflow: hidden;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        tbody tr:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
        }

        tbody td {
            display: block;
            padding: 0;
            border: none;
        }

        /* First row - Main info */
        tbody td:nth-child(1),
        tbody td:nth-child(2),
        tbody td:nth-child(3) {
            display: inline-block;
            vertical-align: top;
        }

        tbody td:nth-child(1) {
            width: 40%;
            padding: 20px 20px 10px 20px;
            font-size: 1.1em;
            font-weight: 600;
        }

        tbody td:nth-child(2) {
            width: 20%;
            padding: 20px 10px 10px 10px;
            text-align: left;
        }

        tbody td:nth-child(3) {
            width: 40%;
            padding: 20px 20px 10px 10px;
            color: #718096;
            font-size: 0.9em;
        }

        /* Second row - Details grid */
        tbody td:nth-child(4),
        tbody td:nth-child(5),
        tbody td:nth-child(9) {
            display: inline-block;
            width: 33.33%;
            padding: 10px 20px;
            border-top: 1px solid #e2e8f0;
            vertical-align: top;
            text-align: left;
        }

        /* Third row - Additional info */
        tbody td:nth-child(6),
        tbody td:nth-child(7),
        tbody td:nth-child(8),
        tbody td:nth-child(10),
        tbody td:nth-child(11),
        tbody td:nth-child(12),
        tbody td:nth-child(13) {
            padding: 8px 20px;
            border-top: 1px solid #e2e8f0;
            font-size: 0.85em;
            color: #718096;
        }

        tbody td:before {
            content: attr(data-label);
            font-weight: 600;
            color: #4a5568;
            display: block;
            margin-bottom: 5px;
            font-size: 0.75em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        tbody td:nth-child(1):before { content: "Domain"; }
        tbody td:nth-child(2):before { content: "MFA Support"; }
        tbody td:nth-child(3):before { content: "Email Address"; }
        tbody td:nth-child(4):before { content: "Documentation"; }
        tbody td:nth-child(5):before { content: "MFA Methods"; }
        tbody td:nth-child(6):before { content: "Custom Software"; }
        tbody td:nth-child(7):before { content: "Custom Hardware"; }
        tbody td:nth-child(8):before { content: "Recovery"; }
        tbody td:nth-child(9):before { content: "Contact"; }
        tbody td:nth-child(10):before { content: "Regions"; }
        tbody td:nth-child(11):before { content: "Additional Domains"; }
        tbody td:nth-child(12):before { content: "Keywords"; }
        tbody td:nth-child(13):before { content: "Notes"; }

        /* Hide empty cells */
        tbody td:empty {
            display: none;
        }

        .mfaEnabled {
            background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85em;
            display: inline-block;
            box-shadow: 0 2px 4px rgba(72, 187, 120, 0.3);
        }

        .mfaDisabled {
            background: linear-gradient(135deg, #f56565 0%, #e53e3e 100%);
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85em;
            display: inline-block;
            box-shadow: 0 2px 4px rgba(245, 101, 101, 0.3);
        }

        a {
            color: #667eea;
            text-decoration: none;
            font-weight: 500;
            transition: color 0.2s;
        }

        a:hover {
            color: #764ba2;
            text-decoration: underline;
        }

        .footer {
            padding: 30px 40px;
            background: #f8f9fa;
            border-top: 1px solid #e2e8f0;
            text-align: center;
            color: #718096;
            font-size: 0.9em;
        }

        .badge {
            display: inline-block;
            padding: 4px 10px;
            background: #edf2f7;
            color: #4a5568;
            border-radius: 12px;
            font-size: 0.85em;
            margin: 2px;
        }

        .domain-logo {
            width: 20px;
            height: 20px;
            margin-right: 8px;
            vertical-align: middle;
            border-radius: 4px;
        }

        .contact-link {
            display: inline-block;
            margin: 2px 4px;
            padding: 4px 8px;
            background: #e6f2ff;
            color: #667eea;
            border-radius: 6px;
            font-size: 0.85em;
            text-decoration: none;
        }

        .contact-link:hover {
            background: #cce5ff;
            text-decoration: none;
        }

        .region-badge {
            display: inline-block;
            padding: 3px 8px;
            background: #fff3cd;
            color: #856404;
            border-radius: 6px;
            font-size: 0.8em;
            margin: 2px;
            font-weight: 500;
        }

        @media (max-width: 768px) {
            .header h1 {
                font-size: 1.8em;
            }

            .info-cards {
                grid-template-columns: 1fr;
            }

            .table-container {
                padding: 20px;
            }

            /* Stack cards vertically on mobile */
            tbody td:nth-child(1),
            tbody td:nth-child(2),
            tbody td:nth-child(3),
            tbody td:nth-child(4),
            tbody td:nth-child(5),
            tbody td:nth-child(9) {
                width: 100% !important;
                display: block !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>MFA-DomainSleuth Report</h1>
            <p>Security Analysis of Email Sender Domains</p>
        </div>

        <div class="info-cards">
            <div class="info-card">
                <strong>Generated</strong>
                <span>$(Get-Date -Format "MMM dd, yyyy HH:mm")</span>
            </div>
            <div class="info-card">
                <strong>Date Range</strong>
                <span>Last $daysToCheck days</span>
            </div>
            <div class="info-card">
                <strong>Email Folder</strong>
                <span>$folderName</span>
            </div>
            <div class="info-card">
                <strong>Data Source</strong>
                <span>Microsoft Graph API</span>
            </div>
        </div>

        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Domain</th>
                        <th>MFA Support</th>
                        <th>Email Address</th>
                        <th>Documentation</th>
                        <th>MFA Methods</th>
                        <th>Custom Software</th>
                        <th>Custom Hardware</th>
                        <th>Recovery</th>
                        <th>Contact Info</th>
                        <th>Regions</th>
                        <th>Additional Domains</th>
                        <th>Keywords</th>
                        <th>Notes</th>
                    </tr>
                </thead>
                <tbody>
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
            $matchedDomain = $senderDomain  # Default to sender domain

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
                # Use the matched domain from the script scope
                $matchedDomain = $script:matchedDomain

                # Get website URL (use url field if available, otherwise construct from domain)
                $websiteURL = if ($siteInfo[1].url) { $siteInfo[1].url } else { "https://$matchedDomain" }

                # Get logo/icon from 2fa.directory
                $logoImg = if ($siteInfo[1].img) {
                    $imgSrc = "https://2fa.directory/img/$($siteInfo[1].img).svg"
                    "<img src='$imgSrc' class='domain-logo' alt='$matchedDomain logo' onerror=""this.style.display='none'"">"
                } else {
                    # Try favicon as fallback
                    $faviconSrc = "https://www.google.com/s2/favicons?domain=$matchedDomain&sz=32"
                    "<img src='$faviconSrc' class='domain-logo' alt='$matchedDomain favicon' onerror=""this.style.display='none'"">"
                }

                $documentationLink = "<a href='$($siteInfo[1].documentation)' target='_blank'>Documentation</a>"
                $tfaMethods = $siteInfo[1].tfa -join ', '
                $customSoftware = $siteInfo[1].'custom-software' -join ', '
                $customHardware = $siteInfo[1].'custom-hardware' -join ', '
                $recoveryURL = $siteInfo[1].recovery
                $additionalDomains = $siteInfo[1].'additional-domains' -join ', '
                $keywords = $siteInfo[1].keywords -join ', '
                $notes = $siteInfo[1].notes

                # Build contact information HTML
                $contactInfo = @()
                if ($siteInfo[1].contact.email) {
                    $contactInfo += "<a href='mailto:$($siteInfo[1].contact.email)' class='contact-link' title='Email Support'>&#128231; Email</a>"
                }
                if ($siteInfo[1].contact.twitter) {
                    $contactInfo += "<a href='https://twitter.com/$($siteInfo[1].contact.twitter)' class='contact-link' target='_blank' title='Twitter Support'>&#128038; Twitter</a>"
                }
                if ($siteInfo[1].contact.facebook) {
                    $contactInfo += "<a href='https://facebook.com/$($siteInfo[1].contact.facebook)' class='contact-link' target='_blank' title='Facebook Support'>&#128211; Facebook</a>"
                }
                if ($siteInfo[1].contact.form) {
                    $contactInfo += "<a href='$($siteInfo[1].contact.form)' class='contact-link' target='_blank' title='Support Form'>&#128221; Form</a>"
                }
                $contactHTML = if ($contactInfo.Count -gt 0) { $contactInfo -join ' ' } else { "" }

                # Build regions HTML
                $regionsHTML = ""
                if ($siteInfo[1].regions) {
                    $regionBadges = $siteInfo[1].regions | ForEach-Object { "<span class='region-badge'>$_</span>" }
                    $regionsHTML = $regionBadges -join ' '
                }
            } else {
                $mfaSupport = "No"
                $mfaClass = "mfaDisabled"
                $SenderEmailAddress = $senderEmail
                # Keep sender domain as matched domain
                $matchedDomain = $senderDomain
                $websiteURL = "https://$matchedDomain"
                $logoImg = ""
                $documentationLink = ""
                $tfaMethods = ""
                $customSoftware = ""
                $customHardware = ""
                $recoveryURL = ""
                $additionalDomains = ""
                $keywords = ""
                $notes = ""
                $contactHTML = ""
                $regionsHTML = ""
            }

            if (($mfaSupport -eq "Yes") -or ($showUnsupported -eq "yes" -and $mfaSupport -eq "No")) {
                $script:htmlOutput += @"
<tr>
    <td>$logoImg<a href="$websiteURL" target="_blank">$matchedDomain</a></td>
    <td><span class='$mfaClass'>$mfaSupport</span></td>
    <td>$SenderEmailAddress</td>
    <td>$documentationLink</td>
    <td>$tfaMethods</td>
    <td>$customSoftware</td>
    <td>$customHardware</td>
    <td><a href="$recoveryURL" target="_blank">$recoveryURL</a></td>
    <td>$contactHTML</td>
    <td>$regionsHTML</td>
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

    # Get all messages using pagination with the provided email address
    $allMessages = @()
    $nextLink = $null

    do {
        if ($nextLink) {
            # Use Invoke-MgGraphRequest for subsequent pages
            $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
            $messages = $response.value
            $nextLink = $response.'@odata.nextLink'
        } else {
            # First batch
            $messages = Get-MgUserMessage -UserId $userEmail -Filter $filter -Top 999 -All
            $nextLink = $null
            break  # -All parameter handles pagination automatically
        }

        if ($messages) {
            $allMessages += $messages
            Write-Host "Fetching more emails... (Currently have $($allMessages.Count))" -ForegroundColor Yellow
        }
    } while ($nextLink)

    if ($allMessages.Count -eq 0) {
        $allMessages = $messages
    }

    $messages = $allMessages
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
                </tbody>
            </table>
        </div>

        <div class="footer">
            <p><strong>Total unique domains found:</strong> $($processedDomains.Count)</p>
            <p style="margin-top: 10px;">Powered by <a href="https://2fa.directory" target="_blank">2FA Directory</a> | Generated using Microsoft Graph API</p>
            <p style="margin-top: 5px;"><a href="https://github.com/danrhodes/MFA-DomainSleuth" target="_blank">GitHub Repository</a></p>
        </div>
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
$openReport = Read-Host "`nWould you like to open the report now? (y/n)"
if ($openReport -match '^y') {
    Start-Process $outputFile
}
