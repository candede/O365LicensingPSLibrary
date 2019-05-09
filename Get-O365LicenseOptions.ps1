<#
.Synopsis
 Report Office 365 user license status
.Description
 Report Office 365 license status for all users in the tenant or only for particular users
.Parameter UserPrincipalName
 UPN of the user. You can add multiple users with comma between UPN values. If you don't specify this parameter, it will report license status of all users in the O365 tenant
.Parameter SaveReportToFile
 Save the output report to a file. By default the output will not be saved
.Example
 Get-O365LicenseOptions.ps1
.Example
 Get-O365LicenseOptions.ps1 -UserPrincipalName user1@contoso.com,user2@contoso.com
.Example
 Get-O365LicenseOptions.ps1 -SaveReportToFile
.Link
 http://www.linkedin.com/in/candede
.Notes
 Author: Can Dedeoglu <candedeoglu@hotmail.com>
 LinkedIn: www.linkedin.com/in/candede
 Date: 8 May 2019
 Version: 1.0
#>

param(
[Parameter(Mandatory=$false, Position=0)]
[array]$UserPrincipalName
,
[Parameter(Mandatory=$false, Position=1)]
[switch]$SaveReportToFile
)

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

$logFilePath = ".\O365_license_status_" + ([datetime]::now).ToString("ddMMMyyyy_HHmmsstt") + ".csv"

if($SaveReportToFile.IsPresent){
    Write-Warning ("Saving output to file {0}" -f $logFilePath)
}

if(![string]::IsNullOrEmpty($UserPrincipalName))
{
    $allUPNs = @($UserPrincipalName)
    $allUsers = $allUPNs | ForEach-Object{Get-MsolUser -UserPrincipalName $_}
}
else {
    $allUsers = Get-MsolUser -All
}

foreach($user in $allUsers)
{
    if($user.isLicensed) {
        $addedAccountSkuIds = @($user.Licenses.AccountSkuId)

        foreach($accountSkuID in $addedAccountSkuIds)
        {
            $enabledLicenses = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Success"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
            $disabledLicenses = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "Disabled"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
            $pendingProvisioning = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "PendingProvisioning"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
            $pendingActivation = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "PendingActivation"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","
            $pendingInput = @(($user.Licenses | Where-Object {$_.AccountSkuId -eq $accountSkuID}).ServiceStatus | Where-Object {$_.ProvisioningStatus -eq "PendingInput"} | ForEach-Object{$_.ServicePlan.ServiceName}) -join ","

            $log = [pscustomobject]@{
                Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Status = "User is Licensed"
                AccountSku = $accountSkuID
                EnabledLicenses = $enabledLicenses
                DisabledLicenses = $disabledLicenses
                PendingProvisioning = $pendingProvisioning
                PendingActivation = $pendingActivation
                PendingInput = $pendingInput
                }
            $log
            if($SaveReportToFile.IsPresent){
                $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
            }
        }
    }
    else {
        $log = [pscustomobject]@{
            Date = ([datetime]::now).ToString('ddMMMyyyy HH:mm:ss')
            DisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Status = "User is NOT Licensed"
            AccountSku = ""
            EnabledLicenses = ""
            DisabledLicenses = ""
            PendingProvisioning = ""
            PendingActivation = ""
            PendingInput = ""
            }
        $log
        if($SaveReportToFile.IsPresent){
            $log | Export-Csv -Path $logFilePath -NoTypeInformation -Append
        }
    }
}