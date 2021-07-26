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
# Script Name:                           tslabwap.ps1
# Author: 	                             Vincent Yim | Senior PFE | vyim@microsoft.com
# Last Update (05/06/2020):              Rodrigo Sorbara | Sr. PFE | rods@microsoft.com 
#                                        Weston Malott | PFE | weston.malott@microsoft.com
#                                        Alexander Dondokov | Sr. PFE | alexd@microsoft.com
#
##########################################################################################################

Clear-Host

$cdname= (Get-ChildItem \\E2K19\c$\labfiles\company*create.txt).Name.Replace("create.txt","")

if ($cdname -eq $null) 
    {
        $cdname = Read-Host "Company file not found on E2K19. `nType the new company name (Example: companyNNNN)"
    }
else
    {
        Read-Host "Found Company name $cdname. Press ENTER if correct, or CTRL+C to break"
    }


$CurrentDate = [datetime]::Now
$Certificate = Get-ChildItem -Path "Cert:\LocalMachine\My" | ? { ($_.Subject -like "*.onelearndns.com*") -AND ($_.NotAfter -ge $CurrentDate) }
$LabProviderDnsSuffix = 'onelearndns.com'
$HttpsPrefix = 'https://'

$OutlookWebAppUrl = $HttpsPrefix + "mail" + $cdname + ".$LabProviderDnsSuffix/"
$AutoDiscoverUrl = $HttpsPrefix + $cdname + ".$LabProviderDnsSuffix/"
$AdfsUrl = $HttpsPrefix + "adfs" + $cdname + ".$LabProviderDnsSuffix/"


if ($Certificate -ne $null)
    {
        $CertificateThumbprint = $Certificate.Thumbprint
        $ServiceAccountCredential = Get-Credential -Message "Enter the credential for the On-Premises Federation Service Account."
        
        Write-Host 'Trying to Publish WAP/OWA/AutoDiscover...' -ForegroundColor Green

        $IsWapAdfsSuccessfullyInstalled = $false
        $IsOwaSuccessfullyInstalled = $false
        $IsAutoDiscoverSuccessfullyInstalled = $false
        $IsAllInstalled = $false

        try
            {
                #WAP
                $CurrentWapApp = 'Web Application Proxy for ADFS'
                Write-Host -NoNewline " ∙ Installing $CurrentWapApp "
                Write-Host $AdfsUrl -ForegroundColor Yellow 
                Install-WebApplicationProxy -FederationServiceTrustCredential $ServiceAccountCredential -CertificateThumbprint $CertificateThumbprint -FederationServiceName ("ADFS" + $cdname + '.' + $LabProviderDnsSuffix) -ErrorAction Stop
                $IsWapAdfsSuccessfullyInstalled = $true

                #OWA
                $CurrentWapApp = 'Exchange Server Outlook Web App'
                Write-Host -NoNewline " ∙ Publishing $CurrentWapApp " 
                Write-Host $OutlookWebAppUrl -ForegroundColor Yellow 
                Add-WebApplicationProxyApplication -BackendServerUrl $OutlookWebAppUrl -ExternalCertificateThumbprint $CertificateThumbprint -ExternalUrl $OutlookWebAppUrl -Name 'Ex2019 OWA' -ExternalPreAuthentication PassThrough -ErrorAction Stop
                $IsOwaSuccessfullyInstalled = $true

                #AutoD
                $CurrentWapApp = 'Exchange Server AutoDiscover'
                Write-Host -NoNewline " ∙ Publishing $CurrentWapApp "
                Write-Host -ForegroundColor Yellow  $AutoDiscoverUrl 
                Add-WebApplicationProxyApplication -BackendServerUrl $AutoDiscoverUrl -ExternalCertificateThumbprint $CertificateThumbprint -ExternalUrl $AutoDiscoverUrl -Name 'Autodiscover2019' -ExternalPreAuthentication PassThrough
                $IsAutoDiscoverSuccessfullyInstalled = $true

                $IsAllInstalled = $true
            }
        catch
            {
                $Message = 'Error: Failed to install/publish "' + $CurrentWapApp + '" , please talk to you instructor :( ' + 'Error details: ' + $Error[0].Exception
                Write-Host $Message -ForegroundColor Red
            }
    }
else
    {
        $Message = 'Error: Cannot find suitable SSL Certificate for ADFS Server, please talk to you instructor :('
        Write-Host $Message -ForegroundColor Red
        $IsWapAdfsSuccessfullyInstalled = $false
    }

if ($IsAllInstalled -eq $true)
    {
        $Message = 'The WAP feature is Successfully installed.'
        Write-Host $Message -ForegroundColor Green
    }
else
    {
        $Message = 'Error: Some or all WAP feature(s) was not installed!'
        Write-Host $Message -ForegroundColor Red

        $Message = 'ADFS WAP Installed: ' + $IsWapAdfsSuccessfullyInstalled
        Write-Host $Message -ForegroundColor Yellow

        $Message = 'OWA Published: ' + $IsOwaSuccessfullyInstalled 
        Write-Host $Message -ForegroundColor Yellow

        $Message = 'Autodiscover Published: ' + $IsAutoDiscoverSuccessfullyInstalled 
        Write-Host $Message -ForegroundColor Yellow
    }


# We are using email domain for autodiscover because 
# 1) Outlook looks for email domain first, so faster resolution and
# 2) wildcard certificate *.onelearndns.com will match companyNNNN.onelearndns.com
#    but wildcard certificate will not match autodiscover.companyNNNN.onelearndns.com
