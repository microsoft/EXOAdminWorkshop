param (
    [Parameter(DontShow)][switch]$BypassCompanyFile
)
# Get company name from file created by LCA
if ($BypassCompanyFile -eq $false) {
    if (Test-Path -Path C:\labfiles\companySettings.json) {
        $companySettings = Get-Content C:\labfiles\companySettings.json | ConvertFrom-Json
        $cdname = $companySettings.CompanyPrefix
     }
    else {
        Write-Warning "companySettings.json file was not found in c:\labfiles. Notify your instructor about this."
        exit
    }
}
# Company file override specified
else {
    Write-Warning "BypassCompanyFile switch was used.  Proceeding with manual value..."
    do {
        Write-Host -NoNewline -ForegroundColor Yellow "Enter your assigned company name (e.g., company12345678): " 
        $cdname = Read-Host
    } While ($cdname -notlike "company*" -or $cdname -like "*.*")
}

Write-Host "Cleaning up all server farms"
appcmd.exe clear config -section:webFarms | Out-Null

Write-Host  "Cleaning up all rewrite rules"
appcmd.exe clear config -section:system.webServer/rewrite/globalRules | Out-Null

Write-Host  "Creating server farm for $cdname.onelearndns.com FQDN"
appcmd.exe set config -section:webFarms /+"[name='$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:webFarms /+"[name='$cdname.onelearndns.com'].[address='$cdname.onelearndns.com']" /commit:apphost | Out-Null

Write-Host  "Creating rewrite rules for $cdname.onelearndns.com FQDN"
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTPS}',pattern='on']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null

Write-Host  "Creating server farm for mail$cdname.onelearndns.com FQDN"
appcmd.exe set config -section:webFarms /+"[name='mail$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:webFarms /+"[name='mail$cdname.onelearndns.com'].[address='mail$cdname.onelearndns.com']" /commit:apphost | Out-Null

Write-Host  "Creating rewrite rules for mail$cdname.onelearndns.com FQDN"
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTPS}',pattern='on']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='mail$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_mail$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://mail$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_mail$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_mail$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_mail$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='mail$cdname.onelearndns.com']" /commit:apphost | Out-Null
appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_mail$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_mail$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://mail$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null

#Write-Host  "Creating server farm for adfs$cdname.onelearndns.com FQDN"
#appcmd.exe set config -section:webFarms /+"[name='adfs$cdname.onelearndns.com']" /commit:apphost | Out-Null
#appcmd.exe set config -section:webFarms /+"[name='adfs$cdname.onelearndns.com'].[address='adfs$cdname.onelearndns.com']" /commit:apphost | Out-Null

#Write-Host  "Creating rewrite rules for adfs$cdname.onelearndns.com FQDN"
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTPS}',pattern='on']" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='adfs$cdname.onelearndns.com']" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance_SSL',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://adfs$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_adfs$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True']" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].match.url:"*"" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /+"[name='ARR_adfs$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].conditions.[input='{HTTP_HOST}',pattern='adfs$cdname.onelearndns.com']" /commit:apphost | Out-Null
#appcmd.exe set config -section:system.webServer/rewrite/globalRules /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.type:"Rewrite"" /"[name='ARR_adfs$cdname.onelearndns.com_loadbalance',patternSyntax='Wildcard',stopProcessing='True'].action.url:"https://adfs$cdname.onelearndns.com/`{R:0`}"" /commit:apphost | Out-Null

Write-Host  "ARR configuration is complete."