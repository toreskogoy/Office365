# This is a collection of various commands regarding licenses in Office 365

#Create Credential Object
$Username = "AdminEmail"
$Password = ConvertTo-SecureString "Password" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($Username, $Password)

#Connect to MS Online Service
Connect-MsolService -Credential $Credential

#Connect to Azure AD
Connect-AzureAD -Credential $Credential

