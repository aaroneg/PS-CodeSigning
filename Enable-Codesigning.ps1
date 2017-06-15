#User Configurable
$CertServerString="server1\EnterpriseRoot"

#Not User Configurable
$mydocs=[environment]::getfolderpath("mydocuments")
$wps="$mydocs\WindowsPowerShell"
$me=Get-AdUser $env:USERNAME
$sam=$me.samaccountname[0]
function RequestCert {
    $Request="
[NewRequest]
Subject=`"CN=$sam`"
KeySpec=2
KeyUsage=0x80

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.3 ; Code Signing
[RequestAttributes]
CertificateTemplate = codesigning"
    $Request|out-file $PSScriptRoot\certreq.inf
    certreq -new $PSScriptRoot\certreq.inf $Psscriptroot\$sam.csr
    $requestID=certreq -submit -config $CertServerString $PSScriptRoot\$sam.csr $PSScriptRoot\$sam.cer
    certreq -accept $PSScriptRoot\$sam.cer
    cd $PSScriptRoot
    gci -Name "$sam.*"|Remove-Item
    gci -Name "*.inf"|Remove-Item
}

if (!(Test-Path $wps)) {
    Write-Output "Creating WindowsPowerShell folder under $mydocs"
    New-Item -ItemType Directory -Path $wps|Out-Null
}

if (!(Test-Path "$wps\signisescript.ps1")) {
    Write-Output "Copying Code Signing ISE addon into place.."
    Copy-Item -Path $PSScriptRoot\_Resources\signisescript.ps1 $wps\
    Write-Output "Adding Script signing add-on to your powershell profile and enabling in this session"
    ". $wps\signisescript.ps1"|Out-File -Append $wps\Microsoft.PowerShellISE_profile.ps1
    . $wps\signisescript.ps1
    '$psISE.CurrentPowerShellTab.AddOnsMenu.submenus.Add(“Sign Script”,{Sign-ISEScript},$null) | Out-Null'|Out-File -Append $wps\Microsoft.PowerShellISE_profile.ps1
    $psISE.CurrentPowerShellTab.AddOnsMenu.submenus.Add(“Sign Script”,{Sign-ISEScript},$null) | Out-Null
}

Write-Output "Checking for code signing cert.."
if(!(gci Cert:\CurrentUser\My -CodeSigningCert)) {
    $title="Acquire Certificate"
    $message="You don't seem to have a code signing certificate. Would you like to create one?"
    $yes=New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continue"
    $no=New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Halt"
    $options=[System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
    $result= $host.ui.PromptForChoice($title,$message,$options,0)
    switch ($result) {
        0 {RequestCert}
        1 {exit}
    }
}
Write-Output "If everything looks ok above, you're good to go. The Add-ons menu now has the option to sign a script."
