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
# Script Name:			         tslabadfs.ps1
# Author: 				 Vincent Yim | Senior PFE | vyim@microsoft.com
# Last Update (05/06/2020):              Rodrigo Sorbara | Sr. PFE | rods@microsoft.com 
#                                        Weston Malott | PFE | weston.malott@microsoft.com
#                                        Alexander Dondokov | Sr. PFE | alexd@microsoft.com
#
##########################################################################################################

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
$FederationServiceDisplayName = $cdname.ToUpper() + " STS"
$FederationServiceName = "ADFS" + $cdname.ToUpper() + ".onelearndns.com"
$Certificate = Get-ChildItem -Path "Cert:\LocalMachine\My" | ? { ($_.Subject -like "*.onelearndns.com*") -AND ($_.NotAfter -ge $CurrentDate) }

if ($Certificate -ne $null)
    {
        $CertificateThumbprint = $Certificate.Thumbprint
        $ServiceAccountCredential = Get-Credential -Message "Enter the credential for the On-Premises Federation Service Account."
        
        Write-Host 'Trying to Install ADFS...' -ForegroundColor Green
        try
            {
                Add-WindowsFeature ADFS-Federation -ErrorAction Stop
                Install-AdfsFarm -OverwriteConfiguration -CertificateThumbprint $CertificateThumbprint -FederationServiceDisplayName $FederationServiceDisplayName -FederationServiceName $FederationServiceName -ServiceAccountCredential $ServiceAccountCredential -ErrorAction Stop
                $IsAdfsSuccessfullyInstalled = $true
            }
        catch
            {
                $Message = 'Error: Failed to install ADFS Server, please talk to you instructor :( ' + 'Error details: ' + $Error[0].Exception
                Write-Host $Message -ForegroundColor Red
                $IsAdfsSuccessfullyInstalled = $false
            }
    }
else
    {
        $Message = 'Error: Cannot find suitable SSL Certificate for ADFS Server, please talk to you instructor :('
        Write-Host $Message -ForegroundColor Red
        $IsAdfsSuccessfullyInstalled = $false
    }

if ($IsAdfsSuccessfullyInstalled -eq $true)
    {
        $Message = 'The ADFS feature is Successfully installed.'
        Write-Host $Message -ForegroundColor Green
    }
else
    {
        $Message = 'The ADFS feature was not installed!'
        Write-Host $Message -ForegroundColor Red
    }
