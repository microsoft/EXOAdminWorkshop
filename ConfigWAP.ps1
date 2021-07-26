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
# Script Name:               ConfigWAP.ps1
# Authors:                   Muris Saab | muriss@microsoft.com, Vincent Yim | Senior PFE | vyim@microsoft.com
# Last Update (11/11/2020):  Rodrigo Sorbara | Sr. PFE | rods@microsoft.com
# Update notes:              Support IP tool lifecycle action (LCA) and create autod records
#
# DNSShell PowerShell module provided by http://dnsshell.codeplex.com
#
##########################################################################################################
param (
    [Parameter(DontShow)][switch]$BypassCompanyFile
)

$testexchange=get-pssnapin *Exchange*
cls
if ($testexchange -eq $Null) {write-host -NoNewline -ForegroundColor Red "You are running Windows Powershell, but this script requires Exchange cmdlets.`nPlease run the Exchange Management Shell (shortcut is in the taskbar). `nPress Enter to Exit..."
    read-host
    exit
}

# Get company name and IP from file created by LCA
if ($BypassCompanyFile -eq $false) {
    if (Test-Path -Path C:\labfiles\companySettings.json) {
        $companySettings = Get-Content C:\labfiles\companySettings.json | ConvertFrom-Json
        $cdname = $companySettings.CompanyPrefix
        Write-Host -ForegroundColor Green "Your assigned company name: $cdname"
        $wapipaddress = $companySettings.PublicIPAddress
        Write-Host -ForegroundColor Green "Your assigned external IP address: $wapipaddress"PFELABS\Administrator
     }
    else {
        Write-Warning "companySettings.json file was not found in c:\labfiles. Notify your instructor about this."
        exit
    }
}
# Company file override specified
else {
    Write-Warning "BypassCompanyFile switch was used.  Proceeding with manual values..."
    do {
        Write-Host -NoNewline -ForegroundColor Yellow "Enter your assigned company name (e.g., company12345678): " 
        $cdname = Read-Host
    } While ($cdname -notlike "company*" -or $cdname -like "*.*")

    Do {
        Write-Host -NoNewline -ForegroundColor Yellow "Enter your assigned public IP address: " 
        $wapipaddress = Read-Host 
        $ipObj = [System.Net.IPAddress]::parse($wapipAddress)
        $isValidIP = [System.Net.IPAddress]::tryparse([string]$wapipAddress, [ref]$ipObj)
    } While (-Not $isValidIP)
}

$fqdnmail = "MAIL" + $cdname + ".onelearndns.com"

Do {Write-host " - Waiting 5 seconds for DC NTDS to become available..."
    start-sleep 5
} Until ((get-service -Name ntds -ComputerName ADFS-DC).status -like "Running")

Import-Module DNSShell

# Create external DNS zone and add records

Write-Host -NoNewline " - Creating your external DNS zone for " ; Write-Host -ForegroundColor White -NoNewline "$cdname.onelearndns.com " ; Write-Host "and required records..."
New-DnsZone -ZoneName "$cdname.onelearndns.com" -ZoneType Primary 
New-DnsRecord -ZoneName "$cdname.onelearndns.com" -RecordType A -IPAddress $wapipaddress 
New-DnsRecord -ZoneName "$cdname.onelearndns.com" -RecordType MX -TargetName $fqdnmail -Preference 10
New-DnsRecord -ZoneName "$cdname.onelearndns.com" -RecordType TXT -Text "v=spf1 ip4:$wapipaddress include:spf.protection.outlook.com ~all"
New-DnsRecord -name "_autodiscover._tcp" -RecordType srv -ZoneName "$cdname.onelearndns.com" -TargetName $fqdnmail -port 443 -Weight 0
New-DnsRecord -name "autodiscover" -ZoneName "$cdname.onelearndns.com" -RecordType A -IPAddress $wapipaddress

# Create internal DNS zone and add records

Write-Host -NoNewline " - Creating your internal DNS zone for " ; Write-Host -ForegroundColor White -NoNewline "onelearndns.com " ; Write-Host "and required records..."
New-DnsZone -ZoneName "onelearndns.com" -ZoneType Primary -Server 192.168.1.11
New-DnsRecord -Name $cdname -ZoneName "onelearndns.com" -RecordType A -IPAddress 192.168.1.4 -Server 192.168.1.11
New-DnsRecord -Name ("mail" + $cdname) -ZoneName "onelearndns.com" -RecordType A -IPAddress 192.168.1.4 -Server 192.168.1.11
New-DnsRecord -Name ("adfs" + $cdname) -ZoneName "onelearndns.com" -RecordType A -IPAddress 192.168.1.11 -Server 192.168.1.11
New-DnsRecord -Name ("autodiscover." + $cdname) -ZoneName "onelearndns.com" -RecordType A -IPAddress 192.168.1.4 -Server 192.168.1.11

$cdname | add-content -path ("\\E2K19\C$\labfiles\" + $cdname + "create.txt")

Get-ReceiveConnector "WAP-EDGE\Default internal receive connector WAP-EDGE" | set-receiveconnector -fqdn $fqdnmail
Get-TransportService|Set-Transportservice -ExternalDNSServers 172.16.0.1,4.2.2.1 -InternalDNSServers 192.168.1.11
get-service adam*|restart-service -force

Write-Host  " - Creating host entries for internal servers"
#add-content c:\windows\system32\drivers\etc\hosts -value ("192.168.1.2`t"+"legacy"+$cdname+".onelearndns.com")
Add-Content c:\windows\system32\drivers\etc\hosts -Value ("192.168.1.11`t"+"adfs"+$cdname+".onelearndns.com")
Start-Sleep -Seconds 1.5
Add-Content c:\windows\system32\drivers\etc\hosts -Value ("192.168.1.4`t"+"mail"+$cdname+".onelearndns.com")
Start-Sleep -Seconds 1.5
Add-Content c:\windows\system32\drivers\etc\hosts -Value ("192.168.1.4`t"+$cdname+".onelearndns.com")
Start-Sleep -Seconds 1.5
Add-Content c:\windows\system32\drivers\etc\hosts -Value ("192.168.1.4`t"+"e2k19.pfelabs.local")

Write-Host  " - Exporting Edge Subscription file" 
try
    {New-EdgeSubscription -filename "\\E2K19\C$\labfiles\edgesubscription.xml" -Force -ErrorAction Stop} 
catch
    {New-EdgeSubscription -filename "\\E2K19\C$\labfiles\edgesubscription.xml" -Force}

Write-Host -ForegroundColor Green "`nDONE!`n" 



<# 
$ipinfo = Invoke-RestMethod http://ipinfo.io/json 
$ipinfo.ip 
$resolvedname=[System.Net.Dns]::GetHostByAddress($ipinfo.ip)
#>