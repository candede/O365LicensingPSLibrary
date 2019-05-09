<#
.Synopsis
 Enable Office 365 user license options
.Description
 Enable Office 365 license options for users from a text file or with given UPN values
.Parameter UserPrincipalName
 UPN of the user. You can add multiple users with comma between UPN values. If you use this parameter then you cannot use FromCSVFilePath or FromTXTFilePath parameters
.Parameter FromCSVFilePath
 Enable Office 365 license options for all the users in the CSV file. CSV file must contain a column with name 'UserPrincipalName' that contains the UPN values of users. If you use this parameter then you cannot use UserPrincipalName or FromTXTFilePath parameters
.Parameter FromTXTFilePath
 Enable Office 365 license options for all the users in the TXT file. TXT file must contain the UPN values of users per line. If you use this parameter then you cannot use UserPrincipalName or FromCSVFilePath parameters
.Parameter EnableLicenses
 Office 365 license options that you want to enable on users. If you want to enable multiple licenses then separate them with comma
.Example
 Enable-O365LicenseOptions.ps1 -UserPrincipalName user1@contoso.com -EnableLicenses TEAMS1,AAD_PREMIUM
.Example
 Enable-O365LicenseOptions.ps1 -FromCSVFilePath .\microsoft_teams_users.csv -EnableLicenses TEAMS1
.Example
 Enable-O365LicenseOptions.ps1 -FromTXTFilePath .\powerapps_users.txt -EnableLicenses POWERAPPS_O365_P2
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
[array]$EnableLicenses
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
        return $_.ToString()
    }
}


$enableLicensesString = @($EnableLicenses) -join ","
$logFilePath = ".\O365_enable_license_" + ([datetime]::now).ToString("ddMMMyyyy_HHmmsstt") + ".csv"

##ENABLE O365 LICENSE OPTIONS
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
            RequestedLicenseOptionsToBeEnabled = ""
            LicenseOptionsToBeEnabled = ""
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
            $licensesCanBeEnabledForThisSkuIDString = ""

            $currentlyDisabledLicenses = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName})
            $licensesCanBeEnabledForThisSkuID = @()

            foreach($license in $EnableLicenses)
            {
                if($currentlyDisabledLicenses -contains $license)
                {
                    $licensesCanBeEnabledForThisSkuID += $license
                }
            }
            $licensesCanBeEnabledForThisSkuIDString = $licensesCanBeEnabledForThisSkuID -join ","

            $newDisabledLicenses = @($currentlyDisabledLicenses | Where-Object {@($licensesCanBeEnabledForThisSkuID) -notcontains $_})
            $newDisabledLicenses = @($newDisabledLicenses | Sort-Object | Get-Unique)

            try{
                $licensesToDisable = New-MsolLicenseOptions -AccountSkuId $accountSkuID -DisabledPlans $newDisabledLicenses -ErrorAction Stop
                Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $licensesToDisable -ErrorAction Stop

                $user = Get-MsolUser -UserPrincipalName $upn
                $enabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Success"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                $disabledLicensesNowString = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
                
                if([string]::IsNullOrEmpty($licensesCanBeEnabledForThisSkuIDString)){
                    $comment = "NoLicenseOptionsToEnable"
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
                    RequestedLicenseOptionsToBeEnabled = $enableLicensesString
                    LicenseOptionsToBeEnabled = $licensesCanBeEnabledForThisSkuIDString
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
                    $exceptionMessage += " Some licenses cannot be enabled alone and needs to be enabled with some other licenses. For example SHAREPOINTWAC_DEVELOPER license cannot be enabled alone; you need to enable SHAREPOINT_S_DEVELOPER,SHAREPOINTWAC_DEVELOPER licenses together. However SHAREPOINT_S_DEVELOPER can be enabled alone without enabling SHAREPOINTWAC_DEVELOPER"
                }
                $log = [pscustomobject]@{
                    Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $upn
                    Status = "User is Licensed"
                    AccountSku = $accountSkuID
                    RequestedLicenseOptionsToBeEnabled = $disableLicensesString
                    LicenseOptionsToBeEnabled = $licensesCanBeDisabledForThisSkuIDString
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
            RequestedLicenseOptionsToBeEnabled = ""
            LicenseOptionsToBeEnabled = ""
            LicensesOptionsEnabledNow = ""
            LicensesOptionsDisabledNow = ""
            Comment = "UserNotLicensed"
            }
        $log
        $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
    }
}