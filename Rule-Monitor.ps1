<#

.SYNOPSIS
Monitor all rules on the Exchange server, sending notifications if a new mailbox rule appears on any clients
that is either forwarding or deleting emails upon receipt.
.DESCRIPTION
Monitor all rules on the Exchange server, sending notifications if a new mailbox rule appears on any clients
that is either forwarding or deleting emails upon receipt.
.PARAMETER CheckNow
Optional switch. Scan through ALL inboxes and instead of reporting changes including delete/forward rules,
report on ANY existing rule that is either forwarding or deleting inbound emails (as well as client rules).

#>

param(
	[switch]$CheckNow = $False
)

# Import the Exchange Management Shell cmdlets.
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


##################################################
#                     TWEAKS                     #
##################################################


# Where to deliver notifications (which email address).
$NotificationsAddress = 'postmaster@thestraightpath.email'
# The address used as a 'source' on email notifications.
$NotificationsSource = 'Client Monitor <client-monitor@thestraightpath.email>'
# Email notification subject line.
$NotificationsSubject = "Summary of User Environment Changes"
# Which mail server to dispatch the notification to/through, and the target relay port.
$NotificationsServer = 'relay.internaldomain.withoutauthentication.com'
$NotificationsServerPort = 25

# CSS styling for notifications, included in the HTML <head> tag for each notification.
$NotificationsStyling = @"
<style>
	body { word-wrap: break-word; font-size: 14px; }
	table { overflow-x: auto; }
	table, th, td { border: 1px solid black; }
	td { text-align: left; padding: 10px; background-color: white; }
	th { background-color: #CCCCCC; text-align: left; padding: 4px 10px; }
	tr:hover, td:hover { background-color: #EDEDED; }
	hr { padding: 0; margin: 10px auto; border: none; }
	h1, h2 { padding: 0; margin: 5px 0; }
	h1 { color: #222222; }
	h2 { color: #560D06; }
	.SummaryText { font-size: 16px; color: black; }
	.FilesList { margin: 10px; background-color: #CCCCCC; color: black; font-size: 10px border-radius: 3px; }
	div.DiffsSection { margin-left: 20px; }
</style>
"@

# Notify the target email address in the below time-range regardless of whether or not changes are detected.
$NotificationsDaily = $True
<# A time range within which notifications will be sent to the target as per the above setting.
 #    The "range" is an amount of minutes as a window for sending the notification.
 #    The "start" is the beginning of the range variable.
 # 
 # Keep in mind, if the script is scheduled to run once every 10 minutes and the Range is set to 60 minutes,
 #    you will receive notifications every time the script runs through those 60 minutes (not recommended).
 #
 # Example: Range: 10, Start: "18:05"
 # If the script runs every 10 minutes (like so */10 -- 00, 10, 20, etc) this will dispatch a daily notification at
 #    around 6:10 PM when the script runs.
 #>
$NotificationsDailyTimeRange = 59
$NotificationsDailyTimeStart = "07:56"

# The filename of the "report" that the script maintains in the current directory, to index changes.
$ReportName = ".\MAIL-RULES-MONITOR-INDEXING.txt"

# When a rule matches, which properties/fields are SELECTED by the script to enter into the HTML table.
$DesiredProperties = @(
	"Name", "Description", "Priority", "Enabled",
	"SupportedByTask", "DeleteMessage", "ForwardTo", "ForwardAsAttachmentTo"
)



##################################################
#                      MAIN                      #
##################################################


# Get all client mailboxes. Perhaps refine this method later.
$inboxes = ((Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox `
	| Select-Object WindowsEmailAddress).WindowsEmailAddress | Select-Object Address) | ForEach-Object { $_.Address }

# If no valid mailboxes were found, error out.		
if($inboxes.Count -le 0) { Write-Host "~~ No mailboxes found on the server. Aborting."; exit 255 }

# "BadPlayers" will hold all the rules to be concerned about.
#    If it's already reported in the last object, don't worry about it (unless the CHECKNOW switch is set).
# BADPLAYERS = {
#   INBOX = {
#     RuleID = HashOfDescription,
#     RuleID = HashOfDescription, ...
#   }
# }
$BadPlayers = @{}
# "AlreadyKnown" will be a simple line-separated file of RuleIdentity fields, basically indexing rules already known.
$AlreadyKnown = @{}

# Fetch the previous information, if it exists.
if(Test-Path $ReportName) {
	$AlreadyKnown = (Get-Content $ReportName | ConvertFrom-Json)
} else {
	Write-Host "~~ Previous report not found. Continuing as if the report is being run for the first time."
	$CheckNow = $True
}


# Iterate the list of inboxes.
Write-Host "-- Scanning inboxes for suspicious rules..."
Write-Host "---- Depending on how many mailboxes there are, this may take a while."
foreach($inbox in $inboxes) {
	# Get the mailbox rules for the target mailbox.
	$ruleset = Get-InboxRule -Mailbox "$inbox"
	$RuleIdentities = @{}
	# Iterate through each rule in the list and attach offending rules to "RuleIdentities.
	foreach($rule in $ruleset) {
		$RuleInfo = @{}
		# If the rule has the following properties set to the corresponding values, be concerned about it.
		if($rule.DeleteMessage -eq $True -Or
			 $rule.ForwardTo -ne $null -Or
			 $rule.SupportedByTask -eq $False -Or
			 $rule.ForwardAsAttachmentTo -ne $null) {
				# This rule is set to trigger, add it to the report.
				## Get the rule's condirional description (CONDITION, not the ACTION).
                $_desc = $rule.Description.ConditionDescriptions | Out-String
                $desc = $_desc | ForEach-Object { $_+"`r`n" }
				$descHash = ""
				## Convert the description into a SHA1 hash.
				([System.Security.Cryptography.HashAlgorithm]::Create("SHA1").ComputeHash(
						[System.Text.Encoding]::UTF8.GetBytes($desc)
					) | ForEach-Object { $descHash += $_.ToString("x2") }
				)
				## Add/Note the fields.
				$RuleInfo.Add("ID", $rule.RuleIdentity.ToString())
				$RuleInfo.Add("HASH", $descHash)
				$RuleIdentities.Add($rule.RuleIdentity.ToString(), $RuleInfo)
		}
	}
	# After all of the mailbox's rules are iterated, take the RuleIdentities array (if length > 0) and put onto BadPlayers.
	if($RuleIdentities.Keys.Count -gt 0) { $BadPlayers.Add($inbox, $RuleIdentities) }
	else { $BadPlayers.Add($inbox, $null) }
}




$NOTIFBODY = "<h1>Suspicious Mailbox Rules</h1>"
$NOTIFBODY += "<p class='SummaryText'>The mailbox rules below were detected across the listed mailboxes.<br />`n"
$NOTIFBODY += "These rules are specifically ones that are <b>forwarding</b> and/or <b>deleting</b> mail on arrival, or are client-side rules specifically.</p>"
if($CheckNow -eq $False) {
	$NOTIFBODY += "<p><u>NOTE</u>: The 'CheckNow' parameter was <i>not</i> set, meaning the below items are"
	$NOTIFBODY += " <b>recent changes</b> since the previous check.</p>"
}
$NOTIFBODY_Rpt = ""
$FlipColors = $True
# Access each inbox name...
Write-Host "-- Iterating through each inbox's rules to compare to the prior report."
foreach($inbox in ($BadPlayers.Keys | Sort-Object)) {
	# Check if there's an existing array of offending ruleID fields, or if it's null...
	if($BadPlayers.$inbox -ne $null) {
		# Iterate through the JSON array of RuleIdentity tags.
		$combinedRules = @()
		foreach($_ruleID in $BadPlayers.$inbox.Keys) {
            $ruleID = $BadPlayers.$inbox.$_ruleID
			$mboxRule = Get-InboxRule -Identity $ruleID.ID -Mailbox $inbox
            $_desc = $mboxRule.Description.ConditionDescriptions | Out-String
            $desc = $_desc | ForEach-Object { $_+"`r`n" }
			$descHash = ""
				([System.Security.Cryptography.HashAlgorithm]::Create("SHA1").ComputeHash(
						[System.Text.Encoding]::UTF8.GetBytes($desc)
					) | ForEach-Object { $descHash += $_.ToString("x2") }
				)
			# Only add to the notification if "checknow" is set, or if the AlreadyKnown index DOES NOT contain the ruleID.
			## OR if the hash of the description has changed.
			$CurrentInboxRuleIDs = @()
			($AlreadyKnown.$inbox | Get-Member -MemberType NoteProperty).Name `
                | ForEach-Object { $CurrentInboxRuleIDs += $AlreadyKnown.$inbox.$_.ID }
			if(($CheckNow -eq $True) -Or
				($CurrentInboxRuleIDs.Contains("$($ruleID.ID)") -eq $False) -Or
				($AlreadyKnown.$inbox."$($ruleID.ID)".HASH -ne $descHash)) {
				# Fetch the rule's properties to add to the notification.
				$rule = $mboxRule | Select-Object $DesiredProperties
                if($DesiredProperties.Contains("Description") -eq $True) {
        	        $rule.Description = ($rule.Description.ConditionDescriptions | Out-String)
                    if($AlreadyKnown.$inbox."$($ruleID.ID)".HASH -ne $descHash -And
                        $AlreadyKnown.$inbox."$($ruleID.ID)".HASH -ne $null) {
                        $rule.Description += "`r`n---MODIFIED---"
                    }
                }
                if($DesiredProperties.Contains("ForwardTo") -eq $True) {
        	        $rule.ForwardTo = ($rule.ForwardTo | ForEach-Object { "$($_), " })
                }
                if($DesiredProperties.Contains("SupportedByTask") -eq $True) {
        	        if($rule.SupportedByTask -eq $False) {
        		        $rule.SupportedByTask = "Client-Side Rule"
        	        } else { $rule.SupportedByTask = "Server-Side Rule" }
                }
                $combinedRules += $rule
			}
		}
		if($combinedRules.Count -gt 0) {
			$tableOut = ($combinedRules | ConvertTo-Json -Depth 1 | ConvertFrom-Json) | ConvertTo-Html -Fragment
			$tableOut = $tableOut -Replace "<td>(,\s*)?</td>", "<td>NULL</td>"
			$tableOut = $tableOut -Replace "<td>Client-Side Rule</td>", "<td style='background-color:#660011;color:white;'>Client-Side Rule</td>"
			$tableOut = $tableOut -Replace "<th>SupportedByTask</th>", "<th>Rule Location</th>"
            $tableOut = $tableOut -Replace "---MODIFIED---", "<br /><br /><span style='font-weight:bold;color:red;font-family:monospace'>---MODIFIED---</span>"
			# Set the background color for the current inbox based on the boolean value of the FlipColors variable.
			$bgColor = if($FlipColors -eq $True) {"FFFFFF"} else {"EDEDED"}
			# If CheckNow was set, show a count of each offending rule next to the email address, otherwise suppress it.
			$countText = if($CheckNow -eq $True) {" -- $($combinedRules.Count)"} else {""}
			# If there was a table added for this inbox, add it to the notifications "Rpt" (wrapper) variable.
			$NOTIFBODY_Rpt += "<hr /><table width='100%' style='border-collapse:collapse;'><tr>"
			$NOTIFBODY_Rpt += "<td width='100%' style='border:none;background-color:#$($bgColor);' bgColor='#$($bgColor)'>"
			$NOTIFBODY_Rpt += "<h2>$($inbox)$($countText)</h2><br />`n<div style='margin-left:20px;'>"
			$NOTIFBODY_Rpt += "$tableOut</div><br /><br /></td></tr></table>`n`n`n"
			$FlipColors = -Not $FlipColors
		}
	}
}

# If the NOTIFBODY_Rpt variable is empty, then there's nothing interesting or valuable to report.
if($NOTIFBODY_Rpt -eq "") {
	# Check if the script is within the daily notifications window (and that daily notifications are set).
	$dateDiff = ((Get-Date) - (Get-Date -Date $NotificationsDailyTimeStart)).TotalMinutes
	if($dateDiff -ge 0 -And $dateDiff -le $NotificationsDailyTimeRange -And $NotificationsDaily -eq $True) {
		Write-Host "---- No changes, but the current time is within the predefined Daily notifications range."
		Write-Host "Sending notification..."
		$NOTIFICATION = "<html><head>$NotificationsStyling</head><body><h1>Nothing Detected</h1>"
		$NOTIFICATION += "<p class='SummaryText'>There were no changes (offending rules) detected across the scanned mailboxes.<br />`n"
		$NOTIFICATION += "This notification is enabled and is dispatched based on the schedule set within the script.</p>"
		$NOTIFICATION += "</body></html>"
		# Send the notification.
		Write-Host "-- Sending notification to '" -NoNewLine
		Write-Host $NotificationsAddress -NoNewLine -ForegroundColor Cyan
		Write-Host "'."
		$NotificationsSubject = "Rule Monitor: Daily Notification"
		Send-MailMessage `
			-From $NotificationsSource -To $NotificationsAddress `
			-Subject $NotificationsSubject -SmtpServer $NotificationsServer `
			-Port $NotificationsServerPort -BodyAsHtml `
			-Body $NOTIFICATION
		if($? -ne $True) { Write-Host "~~~~ Dispatching email notification to '$NotificationsAddress' has failed!" }
	} else {
		if($CheckNow -eq $True) {
			Write-Host "**** No offending rules were detected in the search. No notification will be generated."
		} else {
			Write-Host "**** No new offending rules appeared after this search. No notification will be generated."
		}
	}
} else {
	# Add all the collected info onto the main body of the report.
	$NOTIFBODY += $NOTIFBODY_Rpt
	$NOTIFICATION = "<HTML><HEAD>$NotificationsStyling</HEAD><BODY>$NOTIFBODY</BODY></HTML>"
	# Dispatch the email (or try to)...
	Write-Host "-- Sending notification to '" -NoNewLine
	Write-Host $NotificationsAddress -NoNewLine -ForegroundColor Cyan
	Write-Host "'."
	Send-MailMessage `
		-From $NotificationsSource -To $NotificationsAddress `
		-Subject $NotificationsSubject -SmtpServer $NotificationsServer `
		-Port $NotificationsServerPort -BodyAsHtml `
		-Body $NOTIFICATION
	if($? -ne $True) { Write-Host "~~~~ Dispatching email notification to '$NotificationsAddress' has failed!" }
}


# Overwrite (or freshly write) the new report information.
($BadPlayers | ConvertTo-Json -Compress) | Out-File -FilePath $ReportName