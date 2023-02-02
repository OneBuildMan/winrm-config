param (
    [string]$ComputerName = "servername.domain.com"
)

$FirewallParam = @{ 
    DisplayName = 'Windows Remote Management (HTTPS-In)'
    Direction = 'Inbound' 
    LocalPort = 5986 
    Protocol = 'TCP' 
    Action = 'Allow' 
    Program = 'System' 
} 

$ErrorActionPreference = "Stop"

#
# Check if WinRM is enabled and configured
#
try {
    winrm get winrm/config
    if ($LASTEXITCODE) {
        winrm quickconfig -transport:https
    }
}
catch [System.Management.Automation.CommandNotFoundException] {
    "winrm not found"
    exit
}
catch {
    "An error occurred that could not be resolved."
    exit
}

#
# Locate WinRM listeners and addresses
#

winrm enumerate winrm/config/listener

#
# Display WinRM Firewall rules
#
Get-NetFirewallRule -DisplayGroup "Windows Remote Management" | Get-NetFirewallPortFilter | Format-Table

winrm set winrm/config/service/auth '@{Basic="true"}'

# Validate the service settings
winrm get winrm/config/service/Auth

$Cert = New-SelfSignedCertificate -Subject "CN=`"$ComputerName`"" -TextExtension '2.5.29.37={text}1.3.6.1.5.5.7.3.1'
$CertThumbprint = $Cert.Thumbprint

winrm create winrm/config/Listener?Address=*+Transport=HTTPS "@{Hostname=`"$ComputerName`"; CertificateThumbprint=`"$CertThumbprint`"}" 
if ($LASTEXITCODE) {
    Write-Output "The Listener already exists, we'll recreate it with new generated certificate"
    winrm delete winrm/config/Listener?Address=*+Transport=HTTPS
    winrm create winrm/config/Listener?Address=*+Transport=HTTPS "@{Hostname=`"$ComputerName`"; CertificateThumbprint=`"$CertThumbprint`"}"
}

$ExistingRule = Get-NetFirewallPortFilter | Where-Object -Property LocalPort -EQ $FirewallParam.LocalPort

if(-not $ExistingRule -EQ $null){
    Write-Output "Warning: there is already a rule on port"
    Write-Output $ExistingRule.ToString()
}

New-NetFirewallRule @FirewallParam

