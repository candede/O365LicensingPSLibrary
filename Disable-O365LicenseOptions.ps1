<#
.Synopsis
 Disable Office 365 user license options
.Description
 Disable Office 365 license options for users from a text file or with given UPN values
.Parameter UserPrincipalName
 UPN of the user. You can add multiple users with comma between UPN values. If you use this parameter then you cannot use FromCSVFilePath or FromTXTFilePath parameters
.Parameter FromCSVFilePath
 Disable Office 365 license options for all the users in the CSV file. CSV file must contain a column with name 'UserPrincipalName' that contains the UPN values of users. If you use this parameter then you cannot use UserPrincipalName or FromTXTFilePath parameters
.Parameter FromTXTFilePath
 Disable Office 365 license options for all the users in the TXT file. TXT file must contain the UPN values of users per line. If you use this parameter then you cannot use UserPrincipalName or FromCSVFilePath parameters
.Parameter DisableLicenses
 Office 365 license options that you want to disable on users. If you want to disable multiple licenses then separate them with comma
.Example
 Disable-O365LicenseOptions.ps1 -UserPrincipalName user1@contoso.com -DisableLicenses TEAMS1,AAD_PREMIUM
.Example
 Disable-O365LicenseOptions.ps1 -FromCSVFilePath .\not_microsoft_teams_users.csv -DisableLicenses TEAMS1
.Example
 Disable-O365LicenseOptions.ps1 -FromTXTFilePath .\not_powerapps_users.txt -DisableLicenses POWERAPPS_O365_P2
.Link
 http://www.linkedin.com/in/candede
.Notes
 Author: Can Dedeoglu <candedeoglu@hotmail.com>
 LinkedIn: www.linkedin.com/in/candede
 Date: 8 May 2019
 Version: 1.0
#>

[CmdletBinding(DefaultParameterSetName='UserPrincipalName')]
param(
[Parameter(Mandatory=$true, ParameterSetName = 'UserPrincipalName')]
[array]$UserPrincipalName
,
[Parameter(Mandatory=$true, ParameterSetName = 'FromCSVFilePath')]
[String]$FromCSVFilePath
,
[Parameter(Mandatory=$true, ParameterSetName = 'FromTXTFilePath')]
[String]$FromTXTFilePath
,
[Parameter(Mandatory=$true)]
[array]$DisableLicenses
)

if(![string]::IsNullOrEmpty($UserPrincipalName))
{
    $allUPNs = @($UserPrincipalName)
}

if(![string]::IsNullOrEmpty($FromCSVFilePath))
{
    if(!(Test-Path $FromCSVFilePath))
    {
        return "File not found: $FromCSVFilePath"
    }

    $allCSVData = Import-Csv -Path $FromCSVFilePath

    $csvFileContainsUPNHeader = $allCSVData | Get-Member | Where-Object{$_.Name -eq "UserPrincipalName"}

    if(!$csvFileContainsUPNHeader)
    {
        return "CSV file must contain 'UserPrincipalName' header"
    }

    $allUPNs = $allCSVData | ForEach-Object {$_.UserPrincipalName} | Where-Object{![string]::IsNullOrEmpty($_)}
}

if(![string]::IsNullOrEmpty($FromTXTFilePath))
{
    if(!(Test-Path $FromTXTFilePath))
    {
        return "File not found: $FromTXTFilePath"
    }

    $allUPNs = Get-Content -Path $FromTXTFilePath | Where-Object{![string]::IsNullOrEmpty($_)}
}

##TEST MSOL CONNECTION
try
{
    $company = Get-MsolCompanyInformation -ErrorAction Stop
}
catch
{
    Connect-MsolService

    try
    {
        $company = Get-MsolCompanyInformation -ErrorAction Stop
    }
    catch
    {
        Write-Warning "Connection to MSol Service failed"
        return
    }
}

$disableLicensesString = @($DisableLicenses) -join ","
$logFilePath = ".\O365_disable_license_" + ([datetime]::now).ToString("ddMMMyyyy_HHmmsstt") + ".csv"

##DISABLE O365 LICENSE OPTIONS
foreach($upn in $allUPNs)
{
    $user = ""
    $accountSkuIDs = ""
    try{
        $user = Get-MsolUser -UserPrincipalName $upn -ErrorAction Stop
    }
    catch{
        $log = [pscustomobject]@{
            Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
            DisplayName = ""
            UserPrincipalName = $upn
            Status = "User Not Found"
            AccountSku = ""
            RequestedLicenseOptionsToBeDisabled = ""
            LicenseOptionsToBeDisabled = ""
            LicensesOptionsEnabledNow = ""
            LicensesOptionsDisabledNow = ""
            Comment = "UserNotFound"
            }
        $log
        $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
        return
    }
    
    if($user.isLicensed) {
        $accountSkuIDs = @($user.Licenses.AccountSkuId)
        foreach($accountSkuID in $accountSkuIDs)
        {
            $enabledLicensesNowString = ""
            $disabledLicensesNowString = ""
            $licensesCanBeDisabledForThisSkuIDString = ""

            $allAvailableLicenses = @((Get-MsolAccountSku | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Success"} |ForEach-Object {$_.ServicePlan.ServiceName})
            $currentlyDisabledLicenses = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName})
            $licensesCanBeDisabledForThisSkuID = @()

            foreach($license in $DisableLicenses)
            {
                if($allAvailableLicenses -contains $license -and $currentlyDisabledLicenses -notcontains $license)
                {
                    $licensesCanBeDisabledForThisSkuID += $license
                }
            }
            $licensesCanBeDisabledForThisSkuIDString = $licensesCanBeDisabledForThisSkuID -join ","

            $newDisabledLicenses = @($currentlyDisabledLicenses) + @($licensesCanBeDisabledForThisSkuID)
            $newDisabledLicenses = @($newDisabledLicenses | Sort-Object | Get-Unique)

            try{
                $licensesToDisable = New-MsolLicenseOptions -AccountSkuId $accountSkuID -DisabledPlans $newDisabledLicenses -ErrorAction Stop
                Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $licensesToDisable -ErrorAction Stop

                $user = Get-MsolUser -UserPrincipalName $upn
                $enabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Success"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                $disabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                
                if([string]::IsNullOrEmpty($licensesCanBeDisabledForThisSkuIDString)){
                    $comment = "NoLicenseOptionsToDisable"
                }
                else{
                    $comment = "SUCCESS"
                }
    
                $log = [pscustomobject]@{
                    Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $upn
                    Status = "User is Licensed"
                    AccountSku = $accountSkuID
                    RequestedLicenseOptionsToBeDisabled = $disableLicensesString
                    LicenseOptionsToBeDisabled = $licensesCanBeDisabledForThisSkuIDString
                    LicensesOptionsEnabledNow = $enabledLicensesNowString
                    LicensesOptionsDisabledNow = $disabledLicensesNowString
                    Comment = $comment
                    }
                $log
                $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
            }
            catch{
                $user = Get-MsolUser -UserPrincipalName $upn
                $enabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Success"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                $disabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                
                $exceptionMessage = $_.ToString()
                if($exceptionMessage -match "Unable to assign this license")
                {
                    $exceptionMessage += " Some licenses cannot be disabled alone and needs to be disabled with some other licenses. For example SHAREPOINT_S_DEVELOPER license cannot be disabled alone; you need to disable SHAREPOINT_S_DEVELOPER,SHAREPOINTWAC_DEVELOPER licenses together"
                }
                $log = [pscustomobject]@{
                    Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $upn
                    Status = "User is Licensed"
                    AccountSku = $accountSkuID
                    RequestedLicenseOptionsToBeDisabled = $disableLicensesString
                    LicenseOptionsToBeDisabled = $licensesCanBeDisabledForThisSkuIDString
                    LicensesOptionsEnabledNow = $enabledLicensesNowString
                    LicensesOptionsDisabledNow = $disabledLicensesNowString
                    Comment = $exceptionMessage
                    }
                $log
                $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
            }
        }
    }
    else {
        $log = [pscustomobject]@{
            Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
            DisplayName = $user.DisplayName
            UserPrincipalName = $upn
            Status = "User is NOT Licensed"
            AccountSku = ""
            RequestedLicenseOptionsToBeDisabled = ""
            LicenseOptionsToBeDisabled = ""
            LicensesOptionsEnabledNow = ""
            LicensesOptionsDisabledNow = ""
            Comment = "UserNotLicensed"
            }
        $log
        $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
    }
}