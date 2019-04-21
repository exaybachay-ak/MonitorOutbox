### ---> MonitorOutbox.ps1
### Detect stuck items in outbox and notify user

### ---> Configure environment
if(!(Test-Path -Path C:\temp )){
    New-Item -ItemType Directory -Force -Path C:\temp
}

### ---> You have to run this manually to obtain creds for automatic functionality
if(!(Test-Path -Path C:\temp\securestring.txt )){
    read-host -assecurestring "Please enter your password" | convertfrom-securestring | out-file C:\temp\securestring.txt
}

# Creating variable to detect whether or not a support email has been generated.  
# This acts as a flag to ensure tickets don't get repeatedly generated while the incident is ongoing, and can be reset once connectivity is re-established
$emailsent = 'na'

while(1)
{
	### ---> You have to configure a username or this will not work
	$username = "user@domain.tld"
	$password = cat C:\temp\securestring.txt | convertto-securestring
	$mycreds = new-object -typename System.Management.Automation.PSCredential `
	         -argumentlist $username, $password

	$usermailbox = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name 

	#### --> Configure connection to Outlook API
	Add-type -assembly "Microsoft.Office.Interop.Outlook"
	$Outlook = New-Object -comobject Outlook.Application
	$namespace = $Outlook.GetNamespace("MAPI")

	### ---> Get all folders so we can grab Outboxes 
	$folders = $namespace.Folders

	foreach($folder in $Folders){
		$ob = $folder.Folders.Item("Outbox").Items
		$obItems = $ob.Count

		if($obItems -gt 0){
			if ($emailsent -ne 'True') {
				$wshell = New-Object -ComObject Wscript.Shell
				$wshell.Popup("Outbox item detected!!!",0,"Done",0x1)

				Send-MailMessage -To "user1@domain.tld" -Cc "user2@domain.tld" -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl "Stuck outbox item for $usermailbox" -Port "587" -Body "This is an automatically generated message.<br>Please check on the status of $usermailbox's outbox, as it appears to have stuck emails.<br><b>Your Server Support Bot</b>" -From $username -BodyAsHtml

				#Set flag to true so it doesn't continue sending emails or popups
				$emailsent = 'True'
			}
			else {
				$wshell = New-Object -ComObject Wscript.Shell
				$wshell.Popup("Outbox item detected!!!",0,"Done",0x1)
			}
		}


	    # If there are zero items in the outbox, you are all set.  Undo the flag and move on
	    else {
	        $emailsent = 'False'
	    }

	}

	# Sleep for a predetermined interval.  15 minutes/900 seconds is reasonable in most cases
	start-sleep -seconds 900
}
