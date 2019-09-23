###
# Danny Davis
# Session Example
# Office 365 & Azure <3 PowerShell
# Get a user from SharePoint Online and Azure AD
# Created: 05/31/19
# Modified: 09/21/19
###

Install-Module Microsoft.Online.SharePoint.PowerShell
Install-Module SharePointPnPPowerShellOnline
Install-Module AzureAD

Save-Module msonline -Repository PSGallery -Path "C:\temp"
Save-Module Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -Path "C:\temp"
Save-Module SharePointPnPPowerShellOnline -Repository PSGallery -Path "C:\temp"
Save-Module AzureAD -Repository PSGallery -Path "C:\temp"

Import-Module SharePointPnPPowerShellOnline
Import-Module AzureAD
Import-Module MSOnline

# Azure Tenant Id
# the tenant id is needed to identify the tenant you want to use
# the tenant id can be found within the Azure AD settings
$tenant = ""

# SharePoint Online Central Administration url
# We want to get user information, which is stored in the UPS
# So we have to connect to the CA to get to the UPS and finally retrieve the user data
$urlAdmin = ""

# SharePoint Online Site Collection url
# Get the logins of a group of users who are members / visitors / owners of a Site Collection
# We will get and set the user data of those userss
$url = ""

# We will have to provide credentials to access the cloud
# there are multiple ways, in this case we simply provide them manually
$credentials = Get-Credential

# Connect to all the services
Connect-PnPOnline -Url $urlAdmin -Credentials $credentials
Connect-SPOService -Url $urlAdmin -Credential $credentials
Connect-AzureAD -TenantId $tenant -Credential $credentials

# Get AzureAD User
Get-AzureADUser -ObjectId "user@domain.com"

# Write user to variable
$user = Get-AzureADUser -ObjectId "user@domain.com"

# Display (all) properties of the user
# This list also provides information on what we can change later on
$user 

# Display a specific property of the user
$user.Department

# Set a new job title
# If you try to display the new job title without refreshing your $user, you will not see the information
Set-AzureADUser -ObjectId $user.UserPrincipalName -JobTitle "Marketing"
$user.JobTitle

# Create a new user in Azure and how hopefully not fuck it up
# Create a password profile
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
# Create a super secure default password
$PasswordProfile.Password = "SecurePassword1234!"
# Force user to change password during the next login
$PasswordProfile.ForceChangePasswordNextLogin = $true

# Create a new user in Azure AD
New-AzureADUser -UserPrincipalName "mctest@domain.com" -DisplayName "Testy McTest" -PasswordProfile $PasswordProfile -MailNickName "mctest" -AccountEnabled $true

# With the new user created, we can use the UPN to change user properties
# Save the upn, it's easier to handle from now on
$upn = "mctest@domain.com"
# Set the property "JobTitle"
Set-AzureADUser -ObjectId $UPN -JobTitle "Rockstar"

# We can do the same with the User Profile Server (UPS)
Set-PnPUserProfileProperty -Account $upn -Property "SPS-JobTitle" -Value "Rockstar"