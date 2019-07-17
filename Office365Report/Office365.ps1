#Import Needed module
if (Get-module -ListAvailable -name MSonline) {
    Import-Module MSonline
}
else {
    Write-host "MSOnline module NOT installed. Please run Install-Module MSOnline"
    Break
}

if (Get-module -ListAvailable -name AzureAD) {
    Import-Module AzureAD
}
else {
    Write-Host "AzureAD Module NOT installed. Please run Install-module AzureAD"
}

#Get folder to store report - Shows folder dialog
Function Get-Folder($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if ($foldername.ShowDialog() -eq "OK") {
        $folder += $foldername.SelectedPath
    }
    return $folder
}




############## Set report PATH #################
$ReportPath = Get-Folder
################################################


### Set CSV and Report Name
$LicFile = 'licenses.csv'
$MailFile = 'mailbox.csv'
$GuestFile = 'GuestUsers.csv'
$DeviceFile = 'Devices.csv'
$Outputfile = 'Office365Report.xslx'
$PATH = '\*'

$LICCSV = Join-path $ReportPath -ChildPath $LicFile
$MailCSV = Join-Path $ReportPath -ChildPath $mailfile
$GUESTCSV = Join-Path $ReportPath -ChildPath $GuestFile
$DEVICECSV = Join-path $ReportPath -ChildPath $DeviceFile
$OutputfileName = Join-path $ReportPath -ChildPath $Outputfile
$PATH = Join-path $ReportPath -ChildPath $PATH


#Connecting to Office 365
$Credential = Import-Clixml -Path 'C:\Powershell Scripts\Office 365\servere_no.cred'
Connect-MsolService -Credential $Credential

#Connecting to AzureAD
Connect-AzureAd -Credential $Credential

#Connect to Office 365 Exchange
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection 
Import-PSSession  $ExchangeSession | out-null


#Create list of all products
$Sku = @{
    "O365_BUSINESS_ESSENTIALS"           = "Office 365 Business Essentials"
    "O365_BUSINESS_PREMIUM"              = "Office 365 Business Premium"
    "DESKLESSPACK"                       = "Office 365 (Plan K1)"
    "DESKLESSWOFFPACK"                   = "Office 365 (Plan K2)"
    "LITEPACK"                           = "Office 365 (Plan P1)"
    "EXCHANGESTANDARD"                   = "Office 365 Exchange Online Only"
    "STANDARDPACK"                       = "Enterprise Plan E1"
    "STANDARDWOFFPACK"                   = "Office 365 (Plan E2)"
    "ENTERPRISEPACK"                     = "Enterprise Plan E3"
    "ENTERPRISEPACKLRG"                  = "Enterprise Plan E3"
    "ENTERPRISEWITHSCAL"                 = "Enterprise Plan E4"
    "STANDARDPACK_STUDENT"               = "Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACKPACK_STUDENT"       = "Office 365 (Plan A2) for Students"
    "ENTERPRISEPACK_STUDENT"             = "Office 365 (Plan A3) for Students"
    "ENTERPRISEWITHSCAL_STUDENT"         = "Office 365 (Plan A4) for Students"
    "STANDARDPACK_FACULTY"               = "Office 365 (Plan A1) for Faculty"
    "STANDARDWOFFPACKPACK_FACULTY"       = "Office 365 (Plan A2) for Faculty"
    "ENTERPRISEPACK_FACULTY"             = "Office 365 (Plan A3) for Faculty"
    "ENTERPRISEWITHSCAL_FACULTY"         = "Office 365 (Plan A4) for Faculty"
    "ENTERPRISEPACK_B_PILOT"             = "Office 365 (Enterprise Preview)"
    "STANDARD_B_PILOT"                   = "Office 365 (Small Business Preview)"
    "VISIOCLIENT"                        = "Visio Pro Online"
    "POWER_BI_ADDON"                     = "Office 365 Power BI Addon"
    "POWER_BI_INDIVIDUAL_USE"            = "Power BI Individual User"
    "POWER_BI_STANDALONE"                = "Power BI Stand Alone"
    "POWER_BI_STANDARD"                  = "Power-BI Standard"
    "PROJECTESSENTIALS"                  = "Project Lite"
    "PROJECTCLIENT"                      = "Project Professional"
    "PROJECTONLINE_PLAN_1"               = "Project Online"
    "PROJECTONLINE_PLAN_2"               = "Project Online and PRO"
    "ProjectPremium"                     = "Project Online Premium"
    "ECAL_SERVICES"                      = "ECAL"
    "EMS"                                = "Enterprise Mobility Suite"
    "RIGHTSMANAGEMENT_ADHOC"             = "Windows Azure Rights Management"
    "MCOMEETADV"                         = "PSTN conferencing"
    "SHAREPOINTSTORAGE"                  = "SharePoint storage"
    "PLANNERSTANDALONE"                  = "Planner Standalone"
    "CRMIUR"                             = "CMRIUR"
    "BI_AZURE_P1"                        = "Power BI Reporting and Analytics"
    "INTUNE_A"                           = "Windows Intune Plan A"
    "PROJECTWORKMANAGEMENT"              = "Office 365 Planner Preview"
    "ATP_ENTERPRISE"                     = "Exchange Online Advanced Threat Protection"
    "EQUIVIO_ANALYTICS"                  = "Office 365 Advanced eDiscovery"
    "AAD_BASIC"                          = "Azure Active Directory Basic"
    "RMS_S_ENTERPRISE"                   = "Azure Active Directory Rights Management"
    "AAD_PREMIUM"                        = "Azure Active Directory Premium"
    "MFA_PREMIUM"                        = "Azure Multi-Factor Authentication"
    "STANDARDPACK_GOV"                   = "Microsoft Office 365 (Plan G1) for Government"
    "STANDARDWOFFPACK_GOV"               = "Microsoft Office 365 (Plan G2) for Government"
    "ENTERPRISEPACK_GOV"                 = "Microsoft Office 365 (Plan G3) for Government"
    "ENTERPRISEWITHSCAL_GOV"             = "Microsoft Office 365 (Plan G4) for Government"
    "DESKLESSPACK_GOV"                   = "Microsoft Office 365 (Plan K1) for Government"
    "ESKLESSWOFFPACK_GOV"                = "Microsoft Office 365 (Plan K2) for Government"
    "EXCHANGESTANDARD_GOV"               = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
    "EXCHANGEENTERPRISE_GOV"             = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
    "SHAREPOINTDESKLESS_GOV"             = "SharePoint Online Kiosk"
    "EXCHANGE_S_DESKLESS_GOV"            = "Exchange Kiosk"
    "RMS_S_ENTERPRISE_GOV"               = "Windows Azure Active Directory Rights Management"
    "OFFICESUBSCRIPTION_GOV"             = "Office ProPlus"
    "MCOSTANDARD_GOV"                    = "Lync Plan 2G"
    "SHAREPOINTWAC_GOV"                  = "Office Online for Government"
    "SHAREPOINTENTERPRISE_GOV"           = "SharePoint Plan 2G"
    "EXCHANGE_S_ENTERPRISE_GOV"          = "Exchange Plan 2G"
    "EXCHANGE_S_ARCHIVE_ADDON_GOV"       = "Exchange Online Archiving"
    "EXCHANGE_S_DESKLESS"                = "Exchange Online Kiosk"
    "SHAREPOINTDESKLESS"                 = "SharePoint Online Kiosk"
    "SHAREPOINTWAC"                      = "Office Online"
    "YAMMER_ENTERPRISE"                  = "Yammer Enterprise"
    "EXCHANGE_L_STANDARD"                = "Exchange Online (Plan 1)"
    "MCOLITE"                            = "Lync Online (Plan 1)"
    "SHAREPOINTLITE"                     = "SharePoint Online (Plan 1)"
    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" = "Office ProPlus"
    "EXCHANGE_S_STANDARD_MIDMARKET"      = "Exchange Online (Plan 1)"
    "MCOSTANDARD_MIDMARKET"              = "Lync Online (Plan 1)"
    "SHAREPOINTENTERPRISE_MIDMARKET"     = "SharePoint Online (Plan 1)"
    "OFFICESUBSCRIPTION"                 = "Office ProPlus"
    "YAMMER_MIDSIZE"                     = "Yammer"
    "DYN365_ENTERPRISE_PLAN1"            = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
    "ENTERPRISEPREMIUM_NOPSTNCONF"       = "Enterprise E5 (without Audio Conferencing)"
    "ENTERPRISEPREMIUM"                  = "Enterprise E5 (with Audio Conferencing)"
    "MCOSTANDARD"                        = "Skype for Business Online Standalone Plan 2"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"     = "Dynamics 365 for Financials for IWs"
    "STANDARDWOFFPACK_IW_STUDENT"        = "Office 365 Education for Students"
    "STANDARDWOFFPACK_IW_FACULTY"        = "Office 365 Education for Faculty"
    "EOP_ENTERPRISE_FACULTY"             = "Exchange Online Protection for Faculty"
    "EXCHANGESTANDARD_STUDENT"           = "Exchange Online (Plan 1) for Students"
    "OFFICESUBSCRIPTION_STUDENT"         = "Office ProPlus Student Benefit"
    "STANDARDWOFFPACK_FACULTY"           = "Office 365 Education E1 for Faculty"
    "STANDARDWOFFPACK_STUDENT"           = "Microsoft Office 365 (Plan A2) for Students"
    "DYN365_FINANCIALS_BUSINESS_SKU"     = "Dynamics 365 for Financials Business Edition"
    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU" = "Dynamics 365 for Team Members Business Edition"
    "FLOW_FREE"                          = "Microsoft Flow Free"
    "POWER_BI_PRO"                       = "Power BI Pro"
    "O365_BUSINESS"                      = "Office 365 Business"
    "DYN365_ENTERPRISE_SALES"            = "Dynamics Office 365 Enterprise Sales"
    "RIGHTSMANAGEMENT"                   = "Rights Management"
    "PROJECTPROFESSIONAL"                = "Project Professional"
    "VISIOONLINE_PLAN1"                  = "Visio Online Plan 1"
    "EXCHANGEENTERPRISE"                 = "Exchange Online Plan 2"
    "DYN365_ENTERPRISE_P1_IW"            = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_ENTERPRISE_TEAM_MEMBERS"     = "Dynamics 365 For Team Members Enterprise Edition"
    "CRMSTANDARD"                        = "Microsoft Dynamics CRM Online Professional"
    "EXCHANGEARCHIVE_ADDON"              = "Exchange Online Archiving For Exchange Online"
    "EXCHANGEDESKLESS"                   = "Exchange Online Kiosk"
    "SPZA_IW"                            = "App Connect"
    "WINDOWS_STORE"                      = "Windows Store for Business"
    "MCOEV"                              = "Microsoft Phone System"
    "VIDEO_INTEROP"                      = "Polycom Skype Meeting Video Interop for Skype for Business"
    "SPE_E5"                             = "Microsoft 365 E5"
    "SPE_E3"                             = "Microsoft 365 E3"
    "ATA"                                = "Advanced Threat Analytics"
    "MCOPSTN2"                           = "Domestic and International Calling Plan"
    "FLOW_P1"                            = "Microsoft Flow Plan 1"
    "FLOW_P2"                            = "Microsoft Flow Plan 2"
    "CRMSTORAGE"                         = "Microsoft Dynamics CRM Online Additional Storage"
    "SMB_APPS"                           = "Microsoft Business Apps"
    "MICROSOFT_BUSINESS_CENTER"          = "Microsoft Business Center"
    "DYN365_TEAM_MEMBERS"                = "Dynamics 365 Team Members"
    "STREAM"                             = "Microsoft Stream Trial"
    "EMSPREMIUM"                         = "ENTERPRISE MOBILITY + SECURITY E5"
    "IDENTITY_THREAT_PROTECTION"         = "Identity & Threat Protection"
}

#Create License user class for use with CSV export
class Licuser {
    [string]$Name
    [string]$UserPrincipalname
    [string]$License

    Licuser ([string]$name, [string]$UserPrincipalname, [string]$License) {
        $this.Name = $name
        $this.UserPrincipalname = $UserPrincipalname
        $this.License = $license
    }
    
} 

Class GuestUser {
    [string]$Name
    [string]$UserPrincipalname
    [string]$isLicensed
    [string]$status
    [string]$CreatedDate

    GuestUser ([string]$name, [string]$UserPrincipalname, [string]$isLicensed, [string]$status, [string]$CreatedDate) {
        $this.Name = $name
        $this.UserPrincipalname = $UserPrincipalname
        $this.isLicensed = $isLicensed
        $this.status = $status
        $this.CreatedDate = $CreatedDate

    }

}

Class Device {
    [String]$DeviceName
    [String]$User
    [String]$OS
    [string]$LastLogon

    Device ([String]$DeviceName, [string]$User, [string]$OS, [string]$LastLogon) {
        $this.devicename = $DeviceName
        $this.user = $User
        $this.OS = $OS
        $this.LastLogon = $LastLogon
    }

}

Class Mailbox {
    [string]$Name
    [string]$UserPrincipalname
    [string]$MailboxSize

    Mailbox ([string]$name, [string]$UserPrincipalname, [string]$MailboxSize) {
        $this.name = $name
        $this.UserPrincipalname = $UserPrincipalname
        $this.MailboxSize = $MailboxSize
    }
}

$Users = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq "True" } | Sort-Object DisplayName

foreach ($user in $users) {
    $licenses = ((Get-MsolUser -UserPrincipalName $user.UserPrincipalName).licenses).accountskuid
    if (($licenses).count -gt 0) {
        Foreach ($license in $licenses) {
            Write-host "Working with $($user.displayname)"
            $licenseItem = $license -split ":" | Select-Object -Last 1
            $textlic = $sku.Item("$licenseItem")
            if (!($textlic)) {
                $Lic = [Licuser]::New($User.displayname, $user.UserPrincipalName, $licenseItem)
                $Lic | Export-Csv $LICCSV -NoTypeInformation -Append -Encoding Default
                
                
            }
            else {
                $Lic = [Licuser]::New($User.displayname, $user.UserPrincipalName, $textlic)
                $Lic | Export-Csv $LICCSV -NoTypeInformation -Append -Encoding Default
            }
        }
    }
}
 
#Collect and create Mailbox size Report
$users = Get-Mailbox -ResultSize unlimited | Where-Object { $_.name -ne "DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}" } | Select-Object UserPrincipalName, displayname

If ($users.UserPrincipalName -gt 1) {
    Foreach ($Emailuser in $users) {

        
        #Get mailbox and calculate size into GB
        $mailbox = Get-Mailbox -ResultSize unlimited -Identity $Emailuser.UserPrincipalName | Get-MailboxStatistics
        $data = $mailbox.totalitemsize -replace "(.*\()|,| [a-z]*\)", ""
        $result = [math]::Round($data / 1GB, 2)

        $email = $Emailuser.UserPrincipalName
        $name = $Emailuser.displayname
        $mailboxsize = $result
        $UserObject = [Mailbox]::New($name, $email, $mailboxsize)
        $UserObject | Export-Csv $MAILCSV -NoTypeInformation -Append -Encoding Default
    
    }
}
else {
    Write-Error "No mailboxes found!"
}


#Search for guset user accounts and create report

$GuestUser = Get-AzureADUser -filter "UserType eq 'Guest'" -all $true | Select-Object -property *
Foreach ($Guest in $GuestUser) {
    $name = $Guest.displayname
    $email = $Guest.mail
    $status = $Guest.UserState
    if (!($Guest.AssignedLicenses.skuid)) {
        $license = "No"
    }
    else {
        $license = "Yes"
    }
    $ExtensionProperty = $guest.ExtensionProperty
    $CreatedDate = $ExtensionProperty["createdDateTime"] 

    #Create Object from Class and Subclass
    $GuestUserObject = [GuestUser]::New($Name, $email, $license, $Status, $CreatedDate)
    $GuestUserObject | Export-Csv $GUESTCSV -NoTypeInformation -Append -Encoding Default

}

#Device Report
$Devices = Get-AzureadDevice -all $true | select -Property *
Foreach ($device in $devices) {
    $DeviceName = $device.DisplayName

    Try {
        $UserGuid1 = $device.DevicePhysicalIds -split ":" | select -first 2
        $UserGuid = $UserGuid1[1]
        $User = (Get-AzureADUser -objectid $UserGuid).Displayname
    }
    Catch {
        $User = "Unknown"
    }

    $OS = $device.DeviceOSType
    $LastLogon = $device.ApproximateLastLogonTimeStamp

    $DeviceObject = [Device]::New($DeviceName, $User, $OS, $LastLogon)
    $DeviceObject | Export-Csv $DEVICECSV -NoTypeInformation -Append -Encoding Default
}


#Combine CSV Files into one XLSX report
$CSVS = Get-ChildItem $PATH -Include *.csv
$excel = New-Object -ComObject Excel.Application
$excel.sheetsInNewWorkbook = $CSVS.Count
$xlsx = $excel.Workbooks.add()
$sheet = 1

foreach ($file in $CSVS) {
    $row = 1
    $column = 1
    $worksheet = $xlsx.Worksheets.Item($sheet)
    $worksheet.name = $file.Name
    $filedata = (Get-Content $file)
    foreach ($line in $filedata) {
        $linecontents = $line -split ',(?!\s*\w+")'
        $linecontents = $linecontents.replace("`"", "")
        foreach ($cell in $linecontents ) {
            $worksheet.Cells.Item($row, $column) = $cell
            $column++
        }
        $column = 1
        $row++
    }
    $sheet++
}
$xlsx.SaveAs($outputfilename)
$excel.quit()

#Clean Up Time! Remove CSV files after combined to XLSX file
$LicCSVPath = Test-path ($LICCSV)
If ($LicCSVPath) {
    Remove-Item $LICCSV 
}
$MailCSVPath = Test-Path ($MAILCSV)
if ($MailCSVPath) {
    Remove-item $MAILCSV
}
$GuestCSVpath = Test-Path ($GUESTCSV)
if ($GuestCSVpath) {
    Remove-Item $GUESTCSV
}
$DeviceCSVPath = Test-Path ($DEVICECSV) 
if ($DeviceCSVPath) {
    Remove-Item $DEVICECSV
}


#Remove Exchange PS-Session
Remove-PSSession $ExchangeSession
