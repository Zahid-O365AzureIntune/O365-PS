#Installing MSOnline and AzureAD powershell module
#Install-Module -Name MSOnline
#Install-Module AzureAD

#Setting Execution policy to unrestricted
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

#Storig login credentials
$credential = Get-Credential

#Connect to your Office 365 tenant and AzureAD
#Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All | out-null
Connect-MsolService -Credential $credential
Connect-AzureAD -Credential $credential | out-null

#Newline
Write-Host "`n"

#Sleeps the script for 5s
Sleep -Seconds 5


#Get path of Input .csv file from user
$SourcePath = Read-Host "Enter Input file path"

#Get the Input .csv file name with extension from user
$File = Read-Host "Enter file name"

#Actual location of Input file
$FilePath = Join-Path $SourcePath -ChildPath $File

Write-Output "`n"

#Start capturing the logs
Start-Transcript -Path "$SourcePath\RemoveLicense.log" -Append

#Get csv file contents and store it in a variable
$users=Get-Content $FilePath

#Creating a list to store the available licenses n the tenant
$SkuList = New-Object -TypeName 'System.Collections.ArrayList'


Write-Output "List of The Available Licenses for the tenant based on your Login:"

#Get a list of all the available licenses and display on console
$Sku = Get-MsolAccountSku | select -ExpandProperty AccountSkuId
Write-Output $Sku

Write-Output "`n`n"

#Requesting user to input the AccountSku (License) to be removed
$LicenseToBeRemoved = Read-Host "Enter the license you want to remove"

#Creating a list to store the available licenses n the tenant
$listOfUsers = New-Object -TypeName 'System.Collections.ArrayList'

#Get a list of users with targetted license
foreach ($u in $users)
{
   $userWithLicence= Get-MsolUser -UserPrincipalName $u | Where-Object {($_.licenses).AccountSkuId -match $LicenseToBeRemoved}
   $listOfUsers.Add($userWithLicence) | out-null
}

#Output the UPN with targetted license
Write-Host "Users with the license:"
Write-output $listOfUsers | ft UserPrincipalName

Write-Host "`nRemoving Licenses..."

#Removing targetted license
foreach ($user in $listOfUsers)
{
try
{
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $LicenseToBeRemoved
    Write-Output "Unassigning license for" $user.UserPrincipalName
}

#To handle any errors
Catch
{
  #Write-host $user.UserPrincipalName
  Write-Host "No relevant license assigned"
  #write-host -f Red "No relevant license assigned to "$user.UserPrincipalName"`nError: "$_.Exception.Message
}
}

#End of license removal
Write-Host "Licenses Removed"

#Sleeps the script for 20s
Sleep -Seconds 20

#To disconnect from Azure in your PowerShell session
Disconnect-AzureAD

#End of log capture
Stop-Transcript
