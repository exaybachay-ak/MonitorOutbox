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

#$namespace.Folders.Item(2).Folders.Item("Outbox")

foreach($folder in $Folders){
	$ob = $folder.Folders.Item("Outbox").Items
	$obItems = $ob.Count

	if($obItems -gt 0){
		$wshell = New-Object -ComObject Wscript.Shell
		$wshell.Popup("Outbox item detected!!!",0,"Done",0x1)

		Send-MailMessage -To "user1@domain.tld" -Cc "user2@domain.tld","user3@domain.tld" -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl "Stuck outbox item for $usermailbox" -Port "587" -Body "This is an automatically generated message.<br>Please check on the status of $usermailbox's outbox, as it appears to have stuck emails.<br><b>Your Outbox Monitoring Bot</b>" -From $username -BodyAsHtml
	}
}
