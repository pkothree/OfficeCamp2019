###
# Danny Davis
# Session Example
# Office 365 & Azure <3 PowerShell
# Get a user from SharePoint Online and Azure AD
# Created: 05/31/19
# Modified: 09/21/19
###

# Create Context for PowerShell Modules and User Credentials (connection to O365, O365 Admin)
$FunctionName = 'HttpTriggerPowerShell1'

# Define Modules
$PnPModuleName = 'SharePointPnPPowerShellOnline'
$PnPVersion = '2.24.1803.0'
$AzureADModuleName = 'AzureAD'
$AzureADVersion = '2.0.0.155'
$MSOLModuleName ='MSOnline'
$MSOLVersion ='1.1.166.0'

# Store user in a environmental variable
$username = $Env:user

# Import PS modules
$AzureADModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$AzureADModuleName\$AzureADVersion\$AzureADModuleName.psd1"
$MSOLModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$MSOLModuleName\$MSOLVersion\$MSOLModuleName.psd1"
$PnPModulePath = "D:\home\site\wwwroot\$FunctionName\bin\$PnPModuleName\$PnPVersion\$PnPModuleName.psd1"
$res = "D:\home\site\wwwroot\$FunctionName\bin"
 
Import-Module $AzureADModulePath
Import-Module $PnPModulePath
Import-Module $MSOLModulePath
 
# Build Credentials
$secpassword = ConvertTo-SecureString "Password1234!" -AsPlainText -Force
$credentials= New-Object System.Management.Automation.PSCredential ($username, $secpassword)

# Tenant ID
$tenant = ""

# Connect to MSOL
Connect-MsolService -Credential $credentials

# SharePoint Site Collection URL
$url = ""
$listTitle = "User"

# Connect to SharePoint Online Service
Connect-PnPOnline -Url $url -Credentials $credentials
$item = Get-PNPListItem -List Lists/$listTitle -Id 2

# Connect to Azure AD
Connect-AzureAD -TenantId $tenant -Credential $credentials

if($item.FieldValues.UPN)
{
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $item.FieldValues.Password
    $PasswordProfile.ForceChangePasswordNextLogin = $true

    if($item.FieldValues.MailAddress)
    {
        $split = $item.FieldValues.MailAddress.Split("@")
        $MailNickName = $split[0]
    }

    New-AzureADUser -UserPrincipalName $item.FieldValues.UPN -DisplayName $item.FieldValues.Title -PasswordProfile $PasswordProfile -MailNickName $MailNickName -AccountEnabled $true
}

if($item.FieldValues.Department)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -Department $item.FieldValues.Department.Label
    Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "Department" -Value $item.FieldValues.Department.Label
}
if($item.FieldValues.GivenName)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -GivenName $item.FieldValues.GivenName
}
if($item.FieldValues.SurName)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -Surname $item.FieldValues.SurName
}
if($item.FieldValues.Jobtitle)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -JobTitle $item.FieldValues.Jobtitle.Label
    Set-PnPUserProfileProperty -Account $item.FieldValues.UPN -Property "SPS-JobTitle" -Value $item.FieldValues.Jobtitle.Label
}
if($item.FieldValues.UsageLocation)
{
    Set-AzureADUser -ObjectId $item.FieldValues.UPN -UsageLocation $item.FieldValues.UsageLocation
}