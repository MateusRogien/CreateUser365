<#
Created By: Mariel Borodkin
Created Date: 10/8/2019

.PARAMETER XLSXFilePath
Enter the path of the Excel file containing users to create

.EXAMPLE
CreateUser365.ps1 -XLSXFilePath "c:\temp\Project Users.xlsx"
or
CreateUser365.ps1 "c:\temp\Project Users.xlsx"
#>

#region Param
param(
    [ValidateScript({
        if (-Not ($_ | Test-Path)) {
            throw "File not found"
        }
        elseif (-Not ($_.Extension -eq ".xlsx") ){
            throw "File is not an Excel file (.xlsx)"
        } else {
            return ($true)
        }
    })]
    [System.IO.FileInfo]$XLSXFilePath
)

#endregion Param

#region Functions

<# Function to write logs and outputs to console #>
Function LogWrite{
   Param (
       [Parameter(Mandatory=$true)][string]$logstring,
       [string]$color
   )
   Write-Host $logstring -ForegroundColor $color
   $logstring | Out-File -FilePath $Logfile -Append -Force
}

<# Function to pompt Input box #>
Function InputBox {
    param(

        # Messagebox title
        [Parameter(Mandatory,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $title,

        # Messagebox message
        [Parameter(Mandatory,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $msg
    )

    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    return ($text)
}

<# Function to promt Yes or No box #>
Function YesNoBox {
    param(

        # Messagebox title
        [Parameter(Mandatory,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $title,

        # Messagebox message
        [Parameter(Mandatory,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $msg
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $result = [System.Windows.Forms.MessageBox]::Show($msg , $title , 4)
    if ($result -eq 'Yes') {
        return 1
    } else {
        return 0
    }
}

#endregion .. Functions

#region Variables

<# Global Variables #>

# Logfile
$date = (get-date -Format o).split('.')[0]
$date = $date.Replace('T',' ')
$date = $date.Replace(':','-')

Set-Variable -Name Logfile        -Value "c:\temp\365 Create Time $($date).txt" -Option AllScope
Set-Variable -Name USAGE_LOCATION -Value IL                                     -Option AllScope

<# Constant Variables #>
$ErrorActionPreference = "continue"

Set-Variable -Name LOG_SPLIT           -Value "*********************************************************************" -Option Constant

# Error codes
Set-Variable -Name EXIT_SUCCESS        -Value 0        -Option Constant
Set-Variable -Name EXIT_ERROR          -Value 1        -Option Constant
Set-Variable -Name EXIT_PATH_NOT_FOUND -Value 2        -Option Constant
Set-Variable -Name EXIT_NO_INTERNET    -Value 4        -Option Constant
Set-Variable -name EXIT_UNAUTHORIZED   -Value 5        -Option Constant
Set-Variable -Name EXIT_USER_LEFT      -Value 6        -Option Constant
Set-Variable -Name EXIT_NO_MODULE      -Value 7        -Option Constant

# Action colors
Set-Variable -Name COLOR_ERROR         -Value red      -Option Constant
Set-Variable -Name COLOR_WARNING       -Value yellow   -Option Constant
Set-Variable -Name COLOR_SUCCESS       -Value green    -Option Constant
Set-Variable -Name COLOR_MESSAGE       -Value darkblue -Option Constant

#License groups
Set-Variable -Name LICENSE_OFFICE365   -Value "648b6d87-01af-4b53-8925-96c223929536" -Option Constant
Set-Variable -Name LICENSE_EMSE3       -Value "c0cf3de1-d998-4e36-aa90-8e52bd781157" -Option Constant

$ErrorActionPreference = "stop"
Clear-Host
#endregion .. Variables

#region Log file init
LogWrite $LOG_SPLIT -color $COLOR_MESSAGE
LogWrite "Started Processing at [$($date)]" -color $COLOR_MESSAGE 
LogWrite "$($LOG_SPLIT)" -color $COLOR_MESSAGE
#endregion .. Log file init

#region Prerequisites
    #region Internet Connectivity
if (!(Test-Connection 8.8.8.8 -Count 2 -ErrorAction SilentlyContinue)) {
    LogWrite "ERROR: Internet connection required for script - Exiting script ERROR CODE $($EXIT_NO_INTERNET)" -color $COLOR_MESSAGE
    exit ($EXIT_NO_INTERNET)
}
    #endregion .. Internet Connectivity

    #region Modules installation
if (($null -eq (Get-Module -ListAvailable -Name AzureAD)) -or ($null -eq (Get-Module -ListAvailable -Name MSOnline)) -or ($null -eq (Get-Module -ListAvailable -Name PSExcel))) {
    try {
        Install-Module AzureAD -Confirm:$False -Force
        Install-Module MSOnline -Confirm:$False -Force
        Install-Module PSExcel -Confirm:$False -Force
    } catch {
        Write-Host "Run the following commands in evaluated Powershell Cmdlet: `
        Install-Module AzureAD -Confirm:$False -Force `
        Install-Module MSOnline -Confirm:$False -Force `
        Install-Module PSExcel -Confirm:$False -Force" -ForegroundColor $COLOR_WARNING
    } finally {
        if (($null -eq (Get-Module -ListAvailable -Name AzureAD)) -or ($null -eq (Get-Module -ListAvailable -Name MSOnline)) -or ($null -eq (Get-Module -ListAvailable -Name PSExcel))) {
            LogWrite "** ERROR: You need to install modules before you continue.. Exiting script ERROR CODE $($EXIT_NO_MODULE)" -color $COLOR_ERROR
            exit ($EXIT_NO_MODULE)
        }
    }
    Import-Module PSExcel 
}
    #endregion .. Modules installation
#endregion .. Prerequisites

#region Data - Skipping first line representing the heade
$Users = Import-XLSX $XLSXFilePath
$Users = $Users | Where-Object {$_.id -ne $null}
#endregion .. Data insert

#region Approvals before starting
    #region Users approve
Write-Host "The following steps are required: `
* Approve the table to insert by properties (OK to continue)" -ForegroundColor $COLOR_MESSAGE
$approve = $users | Out-GridView -Title Approval -PassThru
if ($null -eq $approve) {
    LogWrite "** FAILED: Didn't approve the table, please modify the fields before running - Exist Error $($EXIT_USER_LEFT)" -color $COLOR_ERROR
    exit ($EXIT_USER_LEFT)
} else {
    Write-Host "Approved table fields" -ForegroundColor $COLOR_SUCCESS
}
    #endregion .. Users approve

    #region Check open excel 
    $excelProcess = Get-Process -Name *excel*
    if ($null -ne $excelProcess) {
        $allExcel = "You have the following Excel files open:"
        $excelProcess | ForEach-Object {
            $excelname = ($_.MainWindowTitle).Split('-')[0]
            $allExcel = $allExcel + "$($excelname)`n"
        }
        $allExcel = $allExcel + "Please check that the Excel you're using is not open before continue" 
        YesNoBox -title "Validate opened Excel files" -msg $allExcel
    }
    #endregion 

#endregion .. Approvals before starting

#region Authentication
if ((YesNoBox -title "MFA Authentication" -msg "Is your script running account using MFA for authentication?") -eq 0) {
    Write-Host "Insert your credentials`
    1. Azure Active Directory `
    2. Office 365 service `
    3. Exchange Online powershell" -ForegroundColor $COLOR_WARNING
    $UserCredential = Get-Credential
    try {
        Connect-AzureAD -Credential $UserCredential
        Connect-MsolService -Credential $UserCredential
        $a= New-ExoPSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential
        Import-PSSession $a -AllowClobber
    } catch {
        LogWrite "** ERROR: Authentication failed due to wrong password/MFA requirements - Existing script ERROR CODE $($EXIT_UNAUTHORIZED)" -color $COLOR_ERROR
    }
} else { # Case user has MFA and need to insert credentials by his own
    Clear-Host
    Write-Host "Bring your MFA device on! Authentication manually for the following services: `
    1. Azure Active Directory `
    2. Office 365 service `
    3. Exchange Online powershell" -ForegroundColor Yellow
    Connect-AzureAD
    Connect-MsolService
    $a= New-ExoPSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/
    Import-PSSession $a -ErrorAction SilentlyContinue
}

if (($null -eq (Get-AzureADCurrentSessionInfo)) -or ($null -eq (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}))) {
    LogWrite "** ERROR: Authentication for services failed - Existing script ERROR CODE $($EXIT_UNAUTHORIZED)" -color $COLOR_ERROR
    exit ($EXIT_UNAUTHORIZED)
}

#endregion .. Authentication

#region Setting MFA to users
Write-Host "Setting up MFA object.. " -ForegroundColor $COLOR_WARNING
$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enabled"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)
#endregion .. Setting MFA to users

#region Create Users
$AllGroups = Get-AzureADGroup -All $true
LogWrite "------------------------ User creation configuration set ------------------------------" -color $COLOR_MESSAGE

$ErrorActionPreference = "continue"
# Run on each user in Excel
foreach ($User in $Users)
{
    #region User Properties
    $upnName = $User.id -replace '\s',''
    $UPN = $upnName + "@idf.il"
    $fullname = $user.First_Name + " " + $user.Last_Name

    if (($null -ne $user.areacode) -and ($null -ne $user.phone)) {
        $phone = "+$($user.areacode) $($user.phone)"
    } else {
        $phone = $null
    }
    
    if($user.All_Group -eq $null){
        $AllGroup = ($AllGroups | Where-Object {$_.DisplayName -like "*$($user.All_Group)*"}).ObjectId
     }
        $ProjGroup = ($AllGroups | Where-Object {$_.DisplayName -like "*$($user.proj)*"}).ObjectId
    
    #endregion .. User Properties
    
    #region Create Groups
    if ($null -eq $AllGroup -and ($null -ne $user.All_Group)) {
        New-AzureADGroup -DisplayName $user.All_Group -Description "All group $($User.All_Group)" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
        LogWrite "Created All group: $($user.All_Group)" -color $COLOR_MESSAGE
        Start-Sleep -Seconds 5
        $AllGroups = Get-AzureADGroup -All $true
        $AllGroup = ($AllGroups | Where-Object {$_.DisplayName -like "$($user.All_Group)"}).ObjectId
    }

    if ($ProjGroup -like "*-*") {
        New-AzureADGroup -DisplayName $user.Proj -Description -Description "Proj group $($User.Proj)" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
        LogWrite "Created Project group: $($user.Proj)" -color $COLOR_MESSAGE
        Start-Sleep -Seconds 5
        $AllGroups = Get-AzureADGroup -All $true
        $ProjGroup = ($AllGroups | Where-Object {$_.DisplayName -like "$($user.proj)"}).ObjectId
    } else {
        $ProjGroup = ($AllGroups | Where-Object {$_.DisplayName -like "$($user.proj)"}).ObjectId
    }
    #endregion .. Create Groups

    #region Set user

    # Create new user only if not already created
    if ($null -eq (Get-AzureADuser -SearchString $upnName)) {
        New-MsolUser -UserPrincipalName $UPN -DisplayName $fullname -FirstName $user.First_Name -LastName $user.Last_Name -PhoneNumber $phone -MobilePhone $phone -AlternateEmailAddresses $altermail -UsageLocation $USAGE_LOCATION
        Start-Sleep 20
        LogWrite " New User: ID: $($upnName) $($User.First_name) $($User.Last_Name)" -color $COLOR_MESSAGE
    } else {
        LogWrite "Already created: ID: $($upnName) $($User.First_name) $($User.Last_Name)" -color $COLOR_MESSAGE
    }

    # Mechanism to wait if the creation in 365 after AAD takes a little longer than 
    try {
        $msolUser = Get-MsolUser -UserPrincipalName $UPN
    } catch {
        Start-Sleep 10
        $msolUser = Get-MsolUser -UserPrincipalName $UPN
    }
    
    #endregion .. Set user

    #region Define on fields
    if ($User.MFA -like "*yes*") {
        Set-MsolUser -UserPrincipalName $UPN -StrongAuthenticationRequirements $auth
    }
    
    if ($User.OFFICE -like "*yes*") {
        Add-MsolGroupMember -GroupMemberObjectId $msolUser.ObjectId -GroupObjectId $LICENSE_OFFICE365 -ErrorAction SilentlyContinue
    }

    if ($User.EMS -like "*yes*") {
        Add-MsolGroupMember -GroupMemberObjectId $msolUser.ObjectId -GroupObjectId $LICENSE_EMSE3 -ErrorAction SilentlyContinue
    }
    
    if ($ProjGroup -like "*-*") {
        Add-MsolGroupMember -GroupMemberObjectId $msolUser.ObjectId -GroupObjectId $ProjGroup -ErrorAction SilentlyContinue
    }

    Add-MsolGroupMember -GroupMemberObjectId $msolUser.ObjectId -GroupObjectId $AllGroup -ErrorAction SilentlyContinue
    #endregion .. Define on fields
}
#endregion .. Create Users

# Maximal wait time required before creating mail
Start-Sleep -Seconds 480 

#region Change Primary mail

while ($null -eq (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"})) {
    LogWrite "Connection with Exchange online has been lost, authenticate again" -color $COLOR_WARNING
    $a= New-ExoPSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/
    Import-PSSession $a -ErrorAction SilentlyContinue
}

LogWrite "------------------------- Primary Mailbox configuration set -------------------------------" -color $COLOR_MESSAGE

foreach ($User in $Users) {
    $id = $User.id -replace '\s',''
    $altermail = $User.First_Name.split(" ")[0] + $User.Last_Name + "@idf.il"
    $altermail = $altermail -replace '\s',''
    
    try {
        Set-Mailbox -Identity "$id" -EmailAddresses "SMTP:$altermail"
    } catch {
        LogWrite "FAILED: Mail Configuration: $($id) - $ErrorMessage" -color $COLOR_WARNING
    }

    LogWrite "SUCCESFULLY: Mail Configured: $($id)" -color $COLOR_SUCCESS
}
#endregion .. Change Primary mail

Remove-PSSession $a
Start-Process $Logfile
