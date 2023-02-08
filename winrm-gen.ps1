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
# WinRM prechecks
#
try {
    Write-Output "Checking if WinRM is enabled"
    Test-WSMan
    Write-Output "Checking if WinRM configured:"
    winrm get winrm/config
    if ( $LASTEXITCODE ) {
        Write-Output "WinRM not configured, configuring with the default settings"
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

Write-Output "Setting the auth service to Basic"
winrm set winrm/config/service/auth '@{Basic="true"}'
if ( $LASTEXITCODE ) {
    Write-Output "Warning: the auth service could not be set"
}

# Validate the service settings /// winrm set already gives the output.
# Write-Output "Checking the auth service was set"
# winrm get winrm/config/service/Auth

Write-Output "Generating new SSL Certificate"
$Cert = New-SelfSignedCertificate -Subject "CN=`"$ComputerName`"" -TextExtension '2.5.29.37={text}1.3.6.1.5.5.7.3.1'
$CertThumbprint = $Cert.Thumbprint

$ExistingListener = winrm enumerate winrm/config/listener | select-string "Port = 5986"
if( -not ($ExistingListener -EQ $null) ) {
    Write-Output "Warning: there is already a listener on port 5986"
    $Answer = Read-Host -Prompt 'Do you want to create a new one?(The old listener will be deleted) [Y/Any key]'

    if( $Answer -EQ "Y" -or $Answer -EQ "y" ) {
        Write-Output "Deleting old listener"
        winrm delete winrm/config/Listener?Address=*+Transport=HTTPS
        Write-Output "Creating new listener"
        winrm create winrm/config/Listener?Address=*+Transport=HTTPS "@{Hostname=`"$ComputerName`"; CertificateThumbprint=`"$CertThumbprint`"}"
    }
    else {
        Write-Output "Exiting program"
        exit
    }
}
else {
    Write-Output "Creating new listener"
    winrm create winrm/config/Listener?Address=*+Transport=HTTPS "@{Hostname=`"$ComputerName`"; CertificateThumbprint=`"$CertThumbprint`"}"
}

$ExistingRule = Get-NetFirewallPortFilter | Where-Object -Property LocalPort -EQ $FirewallParam.LocalPort
if( -not ($ExistingRule -EQ $null) ) {
    Write-Output "Warning: there is already a rule on port"
}

Write-Output "Opening the appropriate ports on the destination machine’s Windows firewall"
New-NetFirewallRule @FirewallParam

