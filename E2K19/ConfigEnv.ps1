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
# Script Name:		        ConfigEnv.PS1
# Authors:			        Muris Saab | Principal PFE | Microsoft | muris.saab@microsoft.com
# Last Update (2020-07-23):	Rodrigo Sorbara | PFE | rods@microsoft.com (check history in Azure DevOps)
#
##########################################################################################################

cls

#$testexchange=get-pssnapin *Exchange*
#if ($testexchange -eq $Null) {write-host -NoNewline -ForegroundColor Red "You are running the Windows Powershell. Instead, please run Exchange Management Shell from your taskbar. Press Enter to Exit..."
#    read-host
#    exit
#}

$cdname= (dir \\E2K19\c$\labfiles\company*create.txt).name.replace("create.txt","")
if ($cdname -eq $Null) {write-host -NoNewline -ForegroundColor Red "Company file not found on E2K19. `n Rerun configWAP.ps1 on WAPEdgeEx. `n Press Enter to Exit..."
    read-host
    exit
}

$fqdnmail = "mail" + $cdname + ".onelearndns.com"
#$fqdnleg = "legacy" + $cdname + ".onelearndns.com"

Write-Host -NoNewline -ForegroundColor Green "$CDName"
write-host -nonewline -foregroundcolor Yellow " is detected as the company name. `nPress ENTER to provision Exchange servers for $fqdnmail"
Read-Host

# Making sure services are started

Write-Host -ForegroundColor White "`nBEFORE WE BEGIN:"
Write-Host " - Making sure services are started on ADFS-DC..."

    Invoke-Command -ComputerName ADFS-DC -ScriptBlock {Get-WmiObject "Win32_Service"  | where {$_.name -notlike "edge*" -and $_.startmode -eq 'Auto' -and $_.State -eq "Stopped"} | `
	Select-Object Name | Get-Service | Start-Service | Out-Null }

#Write-Host " - Making sure services are started on E2K7..."
#    Get-WmiObject "Win32_Service"  | where {$_.name -notlike "clr_optimization*" -and $_.startmode -eq 'Auto' -and $_.State -eq "Stopped"}  | `
#	Select-Object Name | Get-Service | Start-Service | Out-Null

# Write-Host " - Making sure services are started on E2K10..."
#    Invoke-Command -ComputerName E2K10 -ScriptBlock {Get-WmiObject "Win32_Service"  | where {$_.name -notlike "clr_optimization*" -and $_.startmode -eq 'Auto' -and $_.State -eq "Stopped"}  | `
#	Select-Object Name | Get-Service | Start-Service | Out-Null }

Write-Host " - Making sure services are started on E2K19..."
    Invoke-Command -ComputerName E2K19 -ScriptBlock {Get-WmiObject "Win32_Service"  | where {$_.name -notlike "clr_optimization*" -and $_.startmode -eq 'Auto' -and $_.State -eq "Stopped"} | `
	Select-Object Name | Get-Service | Start-Service | Out-Null }


# Create Remote PowerShell sessions

Write-Host -ForegroundColor White "`nESTABLISHING REMOTE POWERSHELL SESSIONS:"
# Write-Host " - Connecting to E2K10 via Remote PowerShell..."
#     $Session2010 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://E2K10.pfelabs.local/PowerShell
#     Import-PSSession $Session2010 -Prefix 2010 -DisableNameChecking | Out-Null

Write-Host " - Connecting to E2K19 via Remote PowerShell..."
    $Session2019 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://E2K19.pfelabs.local/PowerShell
    Import-PSSession $Session2019 -Prefix 2019 -DisableNameChecking | Out-Null


# Add the new UPN Suffix to Active Directory

Write-Host -ForegroundColor White "`nUPDATING ACTIVE DIRECTORY:"
Write-Host  " - Loading Active Directory PowerShell module..."
    Import-Module ActiveDirectory -ErrorAction:Stop

Write-Host -NoNewline  " - Creating new AD UPN suffix " ; Write-Host -ForegroundColor White "$cdname.onelearndns.com "
    $usn = "$cdname.onelearndns.com"
    $root = [ADSI]"LDAP://rootDSE"
    $conf = [ADSI]"LDAP://cn=partitions,$($root.configurationNamingContext)"
    $conf.uPNSuffixes += $usn
    $conf.SetInfo()

Write-Host  " - Updating the UPN suffix of the Administrator's account..."
    Set-2019Mailbox -Identity Administrator -UserPrincipalName Administrator@$cdname.onelearndns.com

# Configure Exchange
write-host  -ForegroundColor White "`nRecreating Edge Subscription"
New-2019EdgeSubscription -FileData ([byte[]]$(Get-Content -Path "\\E2K19\c$\labfiles\EdgeSubscription.xml" -Encoding Byte -ReadCount 0)) -Site "Default-First-Site-Name"

Write-Host -ForegroundColor White "`nCONFIGURING EXCHANGE WITH YOUR NEW NAMESPACE:"
Write-Host  -NoNewline " - Creating new Accepted Domain " ; Write-Host -NoNewline -ForegroundColor White "$cdname.onelearndns.com " ; Write-Host  "and setting it as default..."
    New-AcceptedDomain -Name "$cdname.onelearndns.com" -DomainType "Authoritative" -DomainName "$cdname.onelearndns.com" | Out-Null
    Set-AcceptedDomain -Identity "$cdname.onelearndns.com" -MakeDefault $true

Write-Host  -NoNewline " - Updating Default Email Address Policy with " ; Write-Host -ForegroundColor White -NoNewline "$cdname.onelearndns.com " ; Write-Host  "as primary SMTP suffix..."
    Set-2019EmailAddressPolicy -Identity "Default Policy" -EnabledEmailAddresstemplates "SMTP:%m@$cdname.onelearndns.com"
    Update-2019EmailAddressPolicy -Identity "Default Policy"
    Remove-AcceptedDomain -Identity pfelabs.local -Confirm:$false

#Write-Host  " - Updating Exchange 2007 URLs..."
#    Set-ClientAccessServer E2K7 -AutoDiscoverServiceInternalUri https://$fqdnmail/Autodiscover/Autodiscover.xml
#    Set-OwaVirtualDirectory "E2K7\owa (Default Web Site)" -ExternalUrl https://$fqdnleg/owa -InternalUrl https://$fqdnleg/owa
#    Set-OabVirtualDirectory "E2K7\oab (Default Web Site)" -ExternalUrl https://$fqdnleg/oab -InternalUrl https://$fqdnleg/oab
#    Set-WebServicesVirtualDirectory "E2K7\ews (Default Web Site)" -ExternalUrl https://$fqdnleg/ews/Exchange.asmx -InternalUrl https://$fqdnleg/ews/Exchange.asmx
#    Set-ActiveSyncVirtualDirectory "E2K7\Microsoft-Server-ActiveSync (Default Web Site)" -ExternalUrl https://$fqdnleg/Microsoft-Server-ActiveSync

# Write-Host  " - Updating Exchange 2010 URLs..."
#     Set-2010ClientAccessServer E2K10 -AutoDiscoverServiceInternalUri https://$fqdnmail/Autodiscover/Autodiscover.xml
#     Set-2010EcpVirtualDirectory "E2K10\ecp (Default Web Site)" -ExternalUrl https://$fqdnmail/ecp -InternalUrl https://$fqdnmail/ecp -WarningAction:SilentlyContinue
#     Set-2010OwaVirtualDirectory "E2K10\owa (Default Web Site)" -ExternalUrl https://$fqdnmail/owa -InternalUrl https://$fqdnmail/owa -WarningAction:SilentlyContinue
#     Set-2010OabVirtualDirectory "E2K10\oab (Default Web Site)" -ExternalUrl https://$fqdnmail/oab -InternalUrl https://$fqdnmail/oab
#     Set-2010WebServicesVirtualDirectory "E2K10\ews (Default Web Site)" -ExternalUrl https://$fqdnmail/ews/Exchange.asmx  -InternalUrl https://$fqdnmail/ews/Exchange.asmx
#     Set-2010ActiveSyncVirtualDirectory "E2K10\Microsoft-Server-ActiveSync (Default Web Site)" -ExternalUrl https://$fqdnmail/Microsoft-Server-ActiveSync

Write-Host  " - Updating Exchange 2019 URLs..."
    Set-2019ClientAccessServer E2K19 -AutoDiscoverServiceInternalUri https://$fqdnmail/Autodiscover/Autodiscover.xml
    Set-2019EcpVirtualDirectory "E2K19\ecp (Default Web Site)" -ExternalUrl https://$fqdnmail/ecp -InternalUrl https://$fqdnmail/ecp -WarningAction:SilentlyContinue
    Set-2019OwaVirtualDirectory "E2K19\owa (Default Web Site)" -ExternalUrl https://$fqdnmail/owa -InternalUrl https://$fqdnmail/owa -WarningAction:SilentlyContinue
    Set-2019OabVirtualDirectory "E2K19\oab (Default Web Site)" -ExternalUrl https://$fqdnmail/oab -InternalUrl https://$fqdnmail/oab
    Set-2019WebServicesVirtualDirectory "E2K19\ews (Default Web Site)" -ExternalUrl https://$fqdnmail/ews/Exchange.asmx  -InternalUrl https://$fqdnmail/ews/Exchange.asmx
    Set-2019ActiveSyncVirtualDirectory "E2K19\Microsoft-Server-ActiveSync (Default Web Site)" -ExternalUrl https://$fqdnmail/Microsoft-Server-ActiveSync
    Set-2019PowerShellVirtualDirectory "E2K19\PowerShell (Default Web Site)" -ExternalUrl https://$fqdnmail/PowerShell

#Write-Host " - Configuring Outlook Anywhere on E2K7..."
#    Get-OutlookAnywhere -Server E2K7 | Set-OutlookAnywhere -ExternalHostname $fqdnmail -ClientAuthenticationMethod Basic -IISAuthenticationMethods Basic, NTLM -SSLOffloading $False -Confirm:$false -WarningAction:SilentlyContinue | Out-Null

# Write-Host " - Configuring Outlook Anywhere on E2K10..."
#     Get-2010OutlookAnywhere -Server E2K10 | Set-2010OutlookAnywhere -ExternalHostname $fqdnmail -ClientAuthenticationMethod NTLM -IISAuthenticationMethods Basic, NTLM -SSLOffloading $False -Confirm:$false -WarningAction:SilentlyContinue | Out-Null

Write-Host " - Configuring Outlook Anywhere on E2K19..."
    Get-2019OutlookAnywhere -Server E2K19 | Set-2019OutlookAnywhere -ExternalHostname $fqdnmail -ExternalClientsRequireSsl $true -InternalHostname $fqdnmail -InternalClientsRequireSsl $true -SSLOffloading $false -ExternalClientAuthenticationMethod NTLM

write-host " - Creating Shortcut on E2K19"
Invoke-Command -ComputerName E2K19 -ScriptBlock {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("c:\users\administrator.pfelabs\desktop\Exchange Admin Center.lnk")
    $Shortcut.TargetPath = "https://mail" + $args[0] + ".onelearndns.com/ecp"
    $shortcut.iconlocation = 'c:\Program Files\Microsoft\Exchange Server\V15\Bin\Exsetupui.exe'
    $Shortcut.Save()
} -argumentlist $cdname

# Create New Mailboxes

Write-Host -ForegroundColor White "`nCREATING SOME MAILBOXES:"
#Write-Host  " - Creating 10 mailboxes on E2K7..."
#11..20 | % {New-Mailbox -Name User$_ -Alias User$_ -DisplayName "User$_ ($cdname)" -Password (ConvertTo-SecureString "Password1" -AsPlainText -Force) -UserPrincipalName "user$_@$cdname.onelearndns.com" -OrganizationalUnit "pfelabs.local/Accounts" -Database (Get-MailboxDatabase -Server E2K7)} | Out-Null

# Write-Host  " - Creating 10 mailboxes on E2K10..."
# 21..30 | % {New-2010Mailbox -Name User$_ -Alias User$_ -DisplayName "User$_ ($cdname)" -Password (ConvertTo-SecureString "Password1" -AsPlainText -Force) -UserPrincipalName "user$_@$cdname.onelearndns.com" -OrganizationalUnit "pfelabs.local/Accounts" -Database (Get-MailboxDatabase -Server E2K10)} | Out-Null
# 21..30 | % {Remove-2010Mailbox -identity User$_ -Confirm:$False} #removing because first set of 2010 mailbox guids can cause conflict issue post-dirsync
# 21..30 | % {New-2010Mailbox -Name User$_ -Alias User$_ -DisplayName "User$_ ($cdname)" -Password (ConvertTo-SecureString "Password1" -AsPlainText -Force) -UserPrincipalName "user$_@$cdname.onelearndns.com" -OrganizationalUnit "pfelabs.local/Accounts" -Database "Mailbox Database 0618600283"} | Out-Null

Write-Host  " - Creating 20 mailboxes on E2K19..."
21..40 | % {New-2019Mailbox -Name User$_ -Alias User$_ -DisplayName "User$_ ($cdname)" -Password (ConvertTo-SecureString "Password1" -AsPlainText -Force) -UserPrincipalName "user$_@$cdname.onelearndns.com" -OrganizationalUnit "pfelabs.local/Accounts" -Database "Mailbox Database 1037417633"} | Out-Null

Write-Host  " - Creating Distribution List for all users..."
New-2019DistributionGroup CompanyDL1 -Type Distribution -DisplayName "CompanyDL1 ($cdname)" | Out-Null
Get-2019Mailbox -RecipientTypeDetails UserMailbox | % {Add-2019DistributionGroupMember CompanyDL1  -Member $_.PrimarySMTPAddress}
# Wrap up

Write-Host -ForegroundColor White "`nWRAPPING UP:"

#Write-Host " - Restarting MSExchangeServiceHost on E2K7..."
#    Restart-Service MSExchangeServiceHost -WarningAction:SilentlyContinue

# Write-Host " - Restarting MSExchangeServiceHost on E2K10..."
#     Invoke-Command -ComputerName E2K10 -ScriptBlock {Restart-Service MSExchangeServiceHost -WarningAction:SilentlyContinue }

Write-Host " - Restarting MSExchangeServiceHost on E2K19..."
    Invoke-Command -ComputerName E2K19 -ScriptBlock {Restart-Service MSExchangeServiceHost -WarningAction:SilentlyContinue }

#Write-Host " - Running IISreset on E2K7..."
#iisreset E2K7 > null

# Write-Host " - Running IISreset on E2K10..."
# iisreset E2K10 > null

Write-Host " - Running IISreset on E2K19..."
iisreset E2K19 > null


Get-PSSession | Remove-PSSession

Write-Host -ForegroundColor Green -NoNewline "`nDone! The environment is now configured with your namespace " ; Write-Host -NoNewline -ForegroundColor White "$cdname.onelearndns.com`n`n"