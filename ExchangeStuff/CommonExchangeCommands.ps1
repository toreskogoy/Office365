$Credential = Get-Credential

#Connect to Office 365 Exchange
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection 
Import-PSSession  $ExchangeSession | out-null

####################################
#        Shared Mailbox            #
####################################

#Add new shared mailbox
New-Mailbox -Name "Post" -Shared -Alias "Post" -PrimarySmtpAddress post@domain.com

#Add permissions to shared mailbox (Use -Autompapping $false in case of no mapping)
Add-MailboxPermission -Identity "post@domain.com" -User "firstname lastname" -AccessRights FullAccess -InheritanceType All

#View Full Access permissons for sbared mailbox
Get-MailboxPermission -Identity "post@domain.com" | Where-Object { ($_.IsInherited -eq $false) -and -not ($_.User -like 'NT AUTHORITY\SELF') } | Select-Object Identity, user, Accessrights

#Remove permissions from shared mailbox
Remove-MailboxPermission -Identity "Post@domain.com" -User "firstname lastname" -AccessRights FullAccess

