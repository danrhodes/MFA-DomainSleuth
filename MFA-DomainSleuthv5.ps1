# Create Outlook COM Object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")

# Choose the folder you want to check in your mailbox
$rootFolder = $namespace.PickFolder()

# Prompt the user for the number of days to check
do {
    $daysToCheck = Read-Host "Enter the number of days to check back on emails"
} while ($daysToCheck -notmatch '^\d+$' -or $daysToCheck -le 0)

do {
    $showUnsupported = Read-Host "Show domains that do not support MFA? (yes/no)"
} while ($showUnsupported -notmatch '^(yes|no)$')


# Create a date object for the specified number of days ago
$startDate = (Get-Date).AddDays(-[int]$daysToCheck)

# Create hashtable for tracking processed domains
$processedDomains = @{}
$matchedDomain = $null

# Get the directory of the current script
$scriptPath = Split-Path -Parent $PSCommandPath

# Fetch JSON data of 2FA enabled/disabled websites
$siteData = Invoke-RestMethod -Uri "https://api.2fa.directory/v3/all.json"

# Initialize HTML output with a modern CSS styling
$script:htmlOutput = @"
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
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
        }
    </style>
</head>
<body>
    <h1>Sender Domains</h1>
    <table>
        <tr>
            <th>Matched Domain</th>
            <th>Supports MFA?</th>
            <th>EMail Address</th>
            <th>Documentation</th>
            <th>TFA Methods</th>
            <th>Custom Software</th>
            <th>Custom Hardware</th>
            <th>Recovery URL</th>
            <th>Additional Domains</th>
            <th>Keywords</th>
            <th>Notes</th>
        </tr>
"@

# Recursive function to process emails in a folder and its subfolders
function ProcessFolder($folder) {
    # Get all email items
    $emails = $folder.Items | Where-Object { $_.ReceivedTime -ge $startDate }

    # Calculate the total number of emails for the progress bar
    $totalEmails = $emails.Count

    # Initialize progress bar
    $i = 0

    # Loop over each email
    foreach ($email in $emails) {
        # Increment progress bar
        $i++
        Write-Progress -Activity "Processing Emails in $($folder.Name)" -Status "$i out of $totalEmails Emails Processed" -PercentComplete (($i / $totalEmails) * 100)
        ProcessEmail $email
    }

    # Recurse into subfolders
    foreach ($subfolder in $folder.Folders) {
        ProcessFolder $subfolder
    }
}

# Function to process each email
function ProcessEmail($email) {
    try {
        # Extract the domain from the sender's email address
        $senderEmailParts = $email.SenderEmailAddress -split "@"
        $senderDomain = $senderEmailParts[-1] -replace ".*@(?<domain>[A-Za-z0-9.-]+(\.[A-Za-z0-9.-]+)+)$", '$1'

        # Skip processing if the domain matches the specified pattern
        if ($senderDomain -like "*EXCHANGELABS*") {
            return
        }

        # Add domain to hashtable and directly write to HTML output if it's not already there
        if (!$processedDomains.ContainsKey($senderDomain)) {
            $processedDomains[$senderDomain] = $true

            # Determine if domain supports MFA and get documentation link
            $siteInfo = $siteData | Where-Object {
                if ($senderDomain -like "*.$($_[1].domain)" -or $senderDomain -eq $_[1].domain) {
                    $matchedDomain = $_[1].domain
                    return $true
                } elseif ($_[1].'additional-domains' | Where-Object { $senderDomain -like "*.$_" -or $senderDomain -eq $_ }) {
                    $matchedDomain = $_
                    return $true
                } else {
                    return $false
                }
            }

            if ($siteInfo) {
                $mfaSupport = "Yes"
                $mfaClass = "mfaEnabled"
                $SenderEmailAddress = $email.SenderEmailAddress
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
                $SenderEmailAddress = $email.SenderEmailAddress
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
    <td><a href="http://$matchedDomain">$matchedDomain</a></td>
    <td class='$mfaClass'>$mfaSupport</td>
    <td>$SenderEmailAddress</td>
    <td>$documentationLink</td>
    <td>$tfaMethods</td>
    <td>$customSoftware</td>
    <td>$customHardware</td>
    <td><a href="$recoveryURL">$recoveryURL</a></td>
    <td>$additionalDomains</td>
    <td>$keywords</td>
    <td>$notes</td>
</tr>
"@
}
        }
    } catch {
        #Write-Host "Error processing email: $($email.Subject) - $_" -ForegroundColor Red
    }
}

# Process the root folder and all its subfolders
ProcessFolder $rootFolder

# Close HTML tags
$script:htmlOutput += "</table></body></html>"

# Write HTML output to file
$script:htmlOutput | Out-File -FilePath "$scriptPath\MFA-DomainSleuth.html"

# Report completion
Write-Host "Completed processing all emails within the last $daysToCheck days. Sender domains have been exported to $scriptPath\MFA-DomainSleuth.html" -ForegroundColor Cyan
