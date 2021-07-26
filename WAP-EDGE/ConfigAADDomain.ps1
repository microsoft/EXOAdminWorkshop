##########################################################################################################
# This sample script is not supported under any Microsoft standard support program or service.
# The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims
# all implied warranties including, without limitation, any implied warranties of merchantability
# or of fitness for a particular purpose. The entire risk arising out of the use or performance of the
# sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone
# else involved in the creation, production, or delivery of the scripts be liable for any damages 
# whatsoever (including, without limitation, damages for loss of business profits, business interruption, 
# loss of business information, or other pecuniary loss) arising out of the use of or inability to use the
# sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages. 
#
# Script Name:               ConfigAADDomain.ps1
# Author:                    Alexander Dondokov | alexd@microsoft.com | Senior PFE 
# Last Update:               ver 0.7 | 01/10/2019 (alexd)
#
##########################################################################################################

#region Variables
$MsOnlineModuleName = 'MsOnline'
$MsOnlineMsiDownloadPage = 'https://www.powershellgallery.com/packages/MSOnline'

$IsConnectedToAAD = $false
$IsDomainNameProvidedResolvable = $false
$DefaultDelayTimeInSecs = 15

$NumberOfLicensesToFreeUp = 5
$AccountSkuToFree = $null
$SkuServicePlanServiceName = 'EXCHANGE_S_ENTERPRISE'
$CompanyAdministratorRoleName = 'Company Administrator'
#endregion

#region Fetching and Installing MSOnline Module if needed
$IsMsOnlineModuleInstalled = [boolean](Get-Module -ListAvailable | ? { $_.Name -eq $MsOnlineModuleName })

if ($IsMsOnlineModuleInstalled -eq $false)
    {
        try
            {
                $Message = 'Installing MsOnline Module...'
                Write-Host $Message -NoNewline
                $MsolPackageInstallationResult = (Find-Package $MsOnlineModuleName | Install-Package -Confirm:$false -Force)
                $IsMsOnlineModuleInstalled = $true   
                Write-Host 'DONE!' -ForegroundColor Green   
            }
        catch
            {
                $Message = 'Cannot install MsOnline Module, please manually download and install it from: "' + $MsOnlineMsiDownloadPage + '".'
                Write-Host $Message -ForegroundColor Red 
                Start-Sleep -Seconds $DefaultDelayTimeInSecs
            }
    }
else
    {
        $IsMsOnlineModuleInstalled = $true
        Import-Module $MsOnlineModuleName
    }
#endregion 

#region User quiz
$MsolCredential = Get-Credential -Message "Please provide Office 365 Tenant Global Admin credential, example: admin@LODSNNNNNN.onmicrosoft.com"
$DomainName = Read-Host "Please provide your assigned domain name, example: companyNNNNNN.onelearndns.com" 

if ( ([string]::IsNullOrWhiteSpace($DomainName)) -ne $true)
    {
        #DNS Resolution
        try
            {
                $DnsResolutionResult = [System.Net.Dns]::GetHostAddresses($DomainName)
                $IsDomainNameProvidedResolvable = $true
            }
        catch
            {
                $IsDomainNameProvidedResolvable = $false
                $Message = 'Error: Failed to resolve domain name provided "' + $DomainName + '".'
                Write-Host $Message -ForegroundColor Red
            }

        #AAD Connection
        try
            {
                Connect-MsolService -Credential $MsolCredential -ErrorAction Stop
                $IsConnectedToAAD = $true
            }
        catch
            {
                $IsConnectedToAAD = $false
                $Message = 'Error: Failed to connect to Azure AD!'
                Write-Host $Message -ForegroundColor Red
            }
    }
else
    {
        $Message = 'Error: The assigned domain name cannot be empty, example: companyNNNNNN.onelearndns.com'
        Write-Host $Message -ForegroundColor Red
        $IsDomainNameProvided = $false
    }
#endregion 

#region Determining if we're eligible to continue or not
if (($IsConnectedToAAD -eq $false) -OR ($IsDomainNameProvidedResolvable -eq $false))
    {
        #Cannot Proceed :( 
        $Message = 'The vital pre-reqs have not been met! Terminating...'
        Write-Host $Message -ForegroundColor Red
        Start-Sleep -Seconds $DefaultDelayTimeInSecs
        Exit
    }
else
    {
        #Do nothing
        $Message = 'All the vital pre-reqs are met! Proceeding...'
        Write-Host $Message -ForegroundColor Green

    }
#endregion 


#region AAD Domain creation
try
    {
        $Message = 'Creating New AAD Domain: "' + $DomainName + '"...'
        Write-Host $Message -NoNewline
        $NewAADDomainCreationResult = New-MsolDomain -Name $DomainName -ErrorAction Stop
        Write-Host 'DONE!' -ForegroundColor Green
    }
catch
    {
        $Message = 'Error: Failed to create the following AAD Domain: "' + $DomainName + '"! ' + $Error[0].Exception
        Write-Host $Message -ForegroundColor Red
    }
#endregion

#region Retrieving Domain Verification DNS Entry
try 
    {
        $MSOLDomainVerificationDnsTxtRecord = Get-MsolDomainVerificationDns -DomainName $DomainName -Mode DnsTxtRecord -ErrorAction Stop
        $Message = 'The Dns TXT verification record was obtained successfully: "' + $MSOLDomainVerificationDnsTxtRecord.Text + '".'
        Write-Host $Message -ForegroundColor Green
    }
catch
    {
        $Message = 'Error: Failed to obtain the DNS Verification record for the following AAD Domain: "' + $DomainName + '"! ' + $Error[0].Exception
        Write-Host $Message -ForegroundColor Red
    }
#endregion

#region Adding Domain Verification TXT record to the External DNS
Add-DnsServerResourceRecord -ZoneName $DomainName -Txt -Name '@' -DescriptiveText $MSOLDomainVerificationDnsTxtRecord.Text -ErrorAction Stop

#Verifying that TXT record has been added successfully
$RetrievedDnsTxtVerificationRecord = Get-DnsServerResourceRecord -ZoneName $DomainName | ? { ($_.RecordType -eq 'TXT') -AND ($_.RecordData.DescriptiveText -eq $MSOLDomainVerificationDnsTxtRecord.Text)}

if ($RetrievedDnsTxtVerificationRecord -ne $null)
    {
        #Sleeping for a few seconds
        $TimeToNapInSecs = $DefaultDelayTimeInSecs

        $Message = 'Executing delay for ' + $TimeToNapInSecs + ' seconds'
        Write-Host $Message -ForegroundColor Green -NoNewline

        While ($TimeToNapInSecs -ne 0)
            {
                $TimeToNapInSecs--
                Write-Host '.' -NoNewline -ForegroundColor (Random(15))
                Start-Sleep -Seconds 1
            }
        Write-Host "`n"

        #Confirm domain in O365 Tenant
        try
            {
                $ConfirmMsolDomainResult = Confirm-MsolDomain -DomainName $DomainName -ErrorAction Stop
                $Message = 'The following domain "' + $DomainName + '" has been confirmed with AAD successfully!'
                Write-Host $Message -ForegroundColor Green
                Get-MsolDomain -DomainName $DomainName
            }
        catch
            {
                $Message = 'Error: Failed to confirm the following AAD Domain: "' + $DomainName + '"! ' + $Error[0].Exception
                Write-Host $Message -ForegroundColor Red
            }
    }
else
    {
        $Message = 'Error: Cannot find the TXT record with the following value: "' + $MSOLDomainVerificationDnsTxtRecord.Text + '".'
        Write-Host $Message -ForegroundColor Red
    }
#endregion 

#region Free-up some licenses
$Message = 'Trying to free up ' + $NumberOfLicensesToFreeUp + ' licenses...'
Write-Host $Message -ForegroundColor Green

$AllMsolAccountSkus = Get-MsolAccountSku

ForEach ($MsolAccountSku in $AllMsolAccountSkus)
    {
        [bool]$IsSkuServicePlanServiceNameFound = (($MsolAccountSku.ServiceStatus.ServicePlan | ? { $_.ServiceName -eq $SkuServicePlanServiceName }) | Measure-Object).Count
        
        if ($IsSkuServicePlanServiceNameFound -eq $true)
            {
                $AccountSkuObjectToFree = $MsolAccountSku
                $AccountSkuToFree = $MsolAccountSku.AccountSkuId
            }
        else
            {
                $AccountSkuToFree = $null
            }
    }

#doing some math here :-)
$NumberOfUnUsedLicenses = $AccountSkuObjectToFree.ActiveUnits - $AccountSkuObjectToFree.ConsumedUnits

if ($NumberOfUnUsedLicenses -ge $NumberOfLicensesToFreeUp)
    {
        [bool]$IsLicenseFreeUpRequired = $false
        $Message = 'No need to free up licenses, there are ' + $NumberOfUnUsedLicenses + ' unused licenses available.'
        Write-Host $Message -ForegroundColor Green

    }
else
    {
        [bool]$IsLicenseFreeUpRequired = $true
        $CalculatedNumberOfLicensesToFreeUp = $NumberOfLicensesToFreeUp - $NumberOfUnUsedLicenses
        $Message = 'There is/are ' + $NumberOfUnUsedLicenses + ' unused licenses available, will free up ' + $CalculatedNumberOfLicensesToFreeUp + ' more.'
        Write-Host $Message -ForegroundColor Green

    }

if (($AccountSkuToFree -ne $null) -AND ($IsLicenseFreeUpRequired -eq $true))
    {
        #FindUsers and FreeUpSomeLicenses
        $CompanyAdministratorRoleObject = Get-MsolRole | ? { $_.Name -eq $CompanyAdministratorRoleName  }
        $CompanyAdministratorObjects = Get-MsolRoleMember -RoleObjectId $CompanyAdministratorRoleObject.ObjectId

        $MsolUserObjectsToUnlicense = New-Object System.Collections.ArrayList
        $AllMsolUserObjects = Get-MsolUser -All
        $AllLicensedMsolUserObjects = $AllMsolUserObjects | ? { $_.IsLicensed -eq $true }
        $AllCloudLicensedMsolUserObjects = $AllLicensedMsolUserObjects | ? { $_.LastDirSyncTime -eq $null }

        ForEach ($CloudLicensedMsolUserObject in $AllCloudLicensedMsolUserObjects)
            {
            
                $IsCloudLicensedMsolUserObjectAdmin = [bool]($CompanyAdministratorObjects.ObjectId | ? { $_ -eq $CloudLicensedMsolUserObject.ObjectId } | Measure-Object).Count
                $IsNecessaryLicenseAssigned = [bool](($CloudLicensedMsolUserObject).Licenses.AccountSkuId | ? { $_ -eq $AccountSkuToFree } | Measure-Object).Count

                if (($IsCloudLicensedMsolUserObjectAdmin -ne $true) -AND ($IsNecessaryLicenseAssigned -eq $true))
                    {
                        $MsolUserObjectsToUnlicense += $CloudLicensedMsolUserObject
                    }
                else
                    {
                        #User is a Company Admin - Skipping...
                        #$CloudLicensedMsolUserObject
                        #$IsCloudLicensedMsolUserObjectAdmin
                    }
            }
        
        $MsolUserObjectsToUnlicenseCandidates = $MsolUserObjectsToUnlicense | Select -Last $CalculatedNumberOfLicensesToFreeUp

        ForEach ($MsolUserObjectsToUnlicenseCandidate in $MsolUserObjectsToUnlicenseCandidates)
            {
                $Message = 'Removing the "' + $AccountSkuToFree + '" license from "' + $MsolUserObjectsToUnlicenseCandidate.UserPrincipalName + '" user...'
                Write-Host $Message -NoNewline
                try
                    {
                        $LicenseRemovalResult = Set-MsolUserLicense -UserPrincipalName $MsolUserObjectsToUnlicenseCandidate.UserPrincipalName -RemoveLicenses $AccountSkuToFree -ErrorAction Stop
                        Write-Host 'DONE!' -ForegroundColor Green
                    }
                catch
                    {
                        Write-Host 'FAILED!' -ForegroundColor Red
                    }
            }
    }
else
    {
        #Warn that licenses must be manually unassigned
    }
#endregion Free-up some licenses

<#region Converting domain to Federated:

The code below only works when the machine is Domain Joined

$AdfsComputerName = 'ADFS3DC'
$AdfsMsolContextLogFile = 'ADFS-Msol.log'

$Message = 'Now switch to the ADFS Server and complete ADFS Setup according to the LAB Manual, then return to the "WapEdgeEx" machine and type "CONTINUE"!'
Write-Host $Message -ForegroundColor DarkRed -BackgroundColor Yellow

while ($ReadHostResult -ne 'CONTINUE')
    {
        $ReadHostResult = Read-Host 'Please type (without quotation marks): "CONTINUE"'
    }

$AdfsUserCredentials = Get-Credential -UserName 'PFELABS\Administrator' -Message "Please provide AD Domain credential of PFELABS\Administrator"

Set-MsolADFSContext -Computer $AdfsComputerName -ADFSUserCredentials $AdfsUserCredentials -LogFile $AdfsMsolContextLogFile 

endregion #>
