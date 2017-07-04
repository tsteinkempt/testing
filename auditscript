# A PowerShell script for auditing the "Audit Log email".

# Scheduling This Script - Windows Task Scheduler

	
		# Syntax
		# PowerShell.exe -command ". '<Path to RemoteExchange.ps1>'; Connect-ExchangeServer -auto; <Path to this Script>"

		# Example
		# PowerShell.exe -command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; C:\Bin\AuditScript2.ps1"


# User Editable Variables

	# Incoming Message Settings

		# Email address the webform is using - e.g. "Help Desk <Helpdesk@nameofmailserver.com>"
		$webformEmailAddress = "Helpdesk@nameofmailserver.com";
		
		# Email addresses of the inboxes that receive submissions
		$reportEmailAddress = "reportrequests@nameofmailserver.com";
		$infoEmailAddress = "info@nameofmailserver.com";
		$mediaEmailAddress = "outreach@nameofmailserver.com";
		$complaintEmailAddress = "complaint.forms@nameofmailserver.com";
		
		# Subject lines of incoming form submissions (note: these are treated as regular expressions by PowerShell)
		$reportSubject = "*Report Request";
		$infoSubject = "*General Information";
		$mediaSubject = "*Media and Events";
		$complaintSubject = "Complaint From*";
		
		# Subject line of the incoming Audit Log email (note: this is treated as a regular expressions by PowerShell)
		$auditSubjectLine = "Office of nameofmailserver Email Logs for the Last 24 Hours *";
		
		# Regular expression for capturing the Audit Log email counts
		$regEx ="\[form_logging__report=(?<report>\d{1,3})\]\[form_logging__info=(?<info>\d{1,3})\]\[form_logging__media=(?<media>\d{1,3})\]\[form_logging__complaint=(?<complaint>\d{1,3})\]"
	
	# Notification Email Settings (local email count is less than the audit report)
	
		# The addresses of people to inform if the number of received submissions does not match the Audit Log
		
		$notificationTo = "DBob@nameofmailserver.com", "BWilliams@nameofmailserver.com", "info@nameofmailserver.com";
		$notificationSubject = "Webform Submissions Error";
		$notificationBody = "One or more Webform submissions have not been received.";
		
	# Receipt Email Settings
	
		$successSubject = "Webform Audit Complete";
		$successBody = "<p> The Audit Script has completed successfully. No discrepencies have been reported. </p>";
		
	# Administrator Email Settings (alerting when something goes wrong)
	
		$adminTo = "Helpdesk@nameofmailserver.com";
		$adminSubject = "Script Error";
		
		# Admin Body Messages (Error Types)
		$errorMailCounting = "<p> The Audit Script was unable to query the Exchange Server Message Tracking Log for one or more submissions. </p>";
		$errorAuditCount = "<p> The Audit Script was unable to capture one or more submission counts from the Audit Log email.</p>";
		$errorTimeSync = "<p> The Audit Script has counted more submissions than reported by the Audit Log email. This is likely due to a time synchronization error.</p>";
		$errorNoAuditLog = "<p> The Audit Script was unable to locate the Audit Log email. This error is usually caused by:</p><ol><li>The Audit Log email failed to arrive before the script was run.</li><li>The Audit Log email has a malformed subject line and is unable to be read as a result.</li></ol>";
		$errorUnknown = "<p> The Audit Script has encountered an error. This is caused by either: </p><ul><li>More than one Audit Log email has been detected.</li><li>An unhandled exception has occured. This will require thorough investigation.</li></ul>";

# Variables

	# Local Counts
	$cReport = $null;
	$cInfo = $null;
	$cMedia = $null;
	$cComplaint = $null;
	
	# Audit Counts
	$aReport = $null;
	$aInfo = $null;
	$aMedia = $null;
	$aComplaint = $null;
	
	# Dates
	$today = ((Get-Date).ToString("MM/dd/yyyy") + " 12:00:00AM");
	$yesterday = ((Get-Date).AddDays(-1).ToString("MM/dd/yyyy") + (" 12:00:00AM"));

# Functions

	# This function counts the number of emails that match a given subject line
	# $subjectLine : the subject line to search for
	# $emailAddress: the recipients to search for
	function countEmails ($subjectLine, $emailAddress) {
		@(Get-MessageTrackingLog -EventID "RECEIVE" -Start $yesterday -End $today -Sender $webformEmailAddress -Recipients $emailAddress -ResultSize unlimited | ? {$_.MessageSubject -like $subjectLine}).Count;
	};

# Script Logic

	# Count the number of submissions for each webform
	$cReport = countEmails $reportSubject $reportEmailAddress;
	$cInfo = countEmails $infoSubject $infoEmailAddress;
	$cMedia = countEmails $mediaSubject $mediaEmailAddress;
	$cComplaint = countEmails $complaintSubject $complaintEmailAddress;
	
	# Check if the script is counting mail correctly
	If (($cReport -eq $null) -or ($cInfo -eq $null) -or ($cMedia -eq $null) -or ($cComplaint -eq $null)) {
		
		# Mail is not being counted correctly, notify administrators
		Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $adminSubject -Body $errorMailCounting -BodyAsHTML -smtpServer nameofmailserver.com;
		Exit;

	};
		
	# Get the Audit Log email
	$auditResults = @(Get-MessageTrackingLog -EventID "RECEIVE" -Start $today -End ((Get-Date).ToString("G")) -Sender $webformEmailAddress -ResultSize unlimited | ? {$_.MessageSubject -like $auditSubjectLine});
	
	# Verify that only one audit log email was found
	If ($auditResults.Count -eq 1) {
		
		# Capture the submission counts from the subject line (piped to supress console output)
		$auditResults[0].MessageSubject -match $regEx | Out-Null;
		
		# Assign the counts
		$aReport = $Matches.report;
		$aInfo = $Matches.info;
		$aMedia = $Matches.media;
		$aComplaint = $Matches.complaint;
		
		# Verify that all webform submission counts have been captured
		If (($aReport -eq $null) -or ($aInfo -eq $null) -or ($aMedia -eq $null) -or ($aComplaint -eq $null)) {
		
			# Unable to capture one or more Audit Log email counts, notify administrators
			Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $adminSubject -Body $errorAuditCount -BodyAsHTML -smtpServer nameofmailserver.com;
			Exit;
		
		};

		# Compare the counts
		If (($cReport -eq $aReport) -and ($cInfo -eq $aInfo) -and ($cMedia -eq $aMedia) -and ($cComplaint -eq $aComplaint)) {
			
			# All counts are correct, send a "success" email and exit
			Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $successSubject -Body $successBody -BodyAsHTML -smtpServer nameofmailserver.com;
			Exit;
		
		# Received less complaints than expected
		} Elseif (($cReport -lt $aReport) -or ($cInfo -lt $aInfo) -or ($cMedia -lt $aMedia) -or ($cComplaint -lt $aComplaint)) {
			
			# Compile error information
			$notificationBody += "<h3>Emails Sent</h3><ul><li>Reports: $aReport</li><li>General Information: $aInfo</li><li>Media and Events: $aMedia</li><li>Complaints: $aComplaint</li></ul><h3>Emails Received</h3><ul><li>Reports: $cReport</li><li>General Information: $cInfo</li><li>Media and Events: $cMedia</li><li>Complaints: $cComplaint</li></ul>";
			
			# Notify the nameofmailserver
			Send-MailMessage -To $notificationTo -From $webformEmailAddress -Subject $notificationSubject -Body $notificationBody -BodyAsHTML -smtpServer nameofmailserver.com;
			Exit;
		
		} Else {
		
			# Received more submissions than reported (likely a time synchronization issue), notify administators
			Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $adminSubject -Body $errorTimeSync -BodyAsHTML -smtpServer nameofmailserver.com;
			Exit;
		
		};
	
	# No Audit Log email was found
	} Elseif ($auditResults -lt 1) {
	
		# Audit email was not received, notify administrators
		Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $adminSubject -Body $errorNoAuditLog -BodyAsHTML -smtpServer nameofmailserver.com;
		Exit;
	
	} Else {

		# Multiple audit emails detected or some other error, notify administrators
		Send-MailMessage -To $adminTo -From $webformEmailAddress -Subject $adminSubject -Body $errorUnknown -BodyAsHTML -smtpServer nameofmailserver.com;
		Exit;

	};

	# Global exit
	Exit;
