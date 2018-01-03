#requires -version 2
<#
.SYNOPSIS
  This script requests a LetsEntryt Certificate for the given DNS name and assigns it to the corresponding IIS site.
.DESCRIPTION
  Using ACMESharp, this scripts requests a certificate for a IIS site, creates an SSL Binding and imports/assigns the certificate.
  It is also capable of scheduling automated renewal jobs.
  Important! Make sure your server is reachable through your domain name via http protocol on port 80! Otherwise it will fail of course.
.INPUTS
  None (Provided in file)
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Marcus Braun
  Creation Date:  3 Jan 2018
  Purpose/Change: Initial script development
#>

#---------------------------------------------------------[Properties]--------------------------------------------------------

$aliasCert = "[MyCertAlias]";
$aliasCertExport = "[MyExportCertNameAlias]";
$dnsEntry = "[DNS Entry e.g. example.com]";

#The path where you will find the certificate (if you dont delete it after finishing the script)
$certStoragePath = "[Path to store cert e.g. c:\vault\]";
$deleteCertificateFileOnCompletion = 1;

#The IIS site name
$SiteName = "[Exact Name of IIS Site]";

#Admin receives a mail prior to expiration
$adminMail = "[Admin Email Adress]";

#Certificate Password
$certificiatePassword = "[SuperSecurePassword]";

# Where to store the imported certificate
$certRootStore = “LocalMachine”
#Do not change the store if not sure what to put here.
$certStore = "My"

#If you have a http to https rewrite rule in place, this tool could automatically disable it before launch. Otherwise set it to 0.
$disableRewriteRule = 1;
$httpsRewriteRuleName = "http to https";

#Add a timestamp to aliases (in case certificate generation fails, you can just rerun the script without much hassle.) 
#Keep in mind, after 20-times of requesting a certificate, LetsEncrypt blocks requests for this DNS Name for 7 days!
#Set 0 if not required.
$addTimestamp = 1;

#do you want the certificate to automatically renew itself after n days? LE certificates last 90 days as of today.
#Set 0 if no required.
$scheduleRenewal = 1;
$scheduleDays = 89;
$scheduleTime = "3am";

#---------------------------------------------------------[Setup]--------------------------------------------------------
$scriptPath = $MyInvocation.ScriptName
try {
    Import-Module ACMESharp
    Write-Host "Imported ACME Sharp..."
} catch {
    Write-Host "Installing ACME Sharp..."
    Install-Module ACMESharp
    Install-Module ACMESharp.Providers.IIS
    Import-Module ACMESharp
    Enable-ACMEExtensionModule ACMESharp.Providers.IIS
}
try {
    Get-Variable acmeReg -ErrorAction Stop
} catch [System.Management.Automation.ItemNotFoundException] {
    $acmeReg = New-ACMERegistration -Contacts mailto:$adminMail -AcceptTos;
}

#Disable potentially active redirects
$redirectEnabled = (Get-WebConfiguration system.webServer/httpRedirect "IIS:\sites\$SiteName").enabled
if(redirectEnabled -eq "True"){
  Set-WebConfiguration system.webServer/httpRedirect "IIS:\sites\$SiteName" -Value @{enabled="false"} -ErrorAction SilentlyContinue
}

if($addTimestamp -eq 1){
  $ts =  (get-date -f MM_dd_yyyy_HH_mm_ss).ToString();

  $aliasCert = ($aliasCert + "_" + $ts);
  $aliasCertExport = ($aliasCertExport + "_" + $ts);

  Write-Host "Adding Timestamp to alias identifiers: " $aliasCert;
  Write-Host "Adding Timestamp to alias export identifiers: " $aliasCertExport;
}


#Generate folder for cert disk storage path if needed
Write-Host "Creating storage vault path if not exists: " $certStoragePath;
md -Force $certStoragePath | Out-Null;

#Removing current SSL Webbinding
Write-Host "Removing SSL Binding...";
Get-WebBinding -Port 443 -Name $SiteName | Remove-WebBinding

#Import ACMEShart Module (https://github.com/ebekker/ACMESharp)
Write-Host "Registering with Letsencrypt...";

New-ACMEIdentifier -Dns $dnsEntry -Alias $aliasCert

#disable url rewrite if needed
if($disableRewriteRule -eq 1){
    $rwStorN = '/system.webserver/rewrite/rules/rule[@name="{0}"]' -f $httpsRewriteRuleName
    Write-Host "Disabling HTTPS Rewrite..."
    set-webconfigurationproperty $rwStorN -Name enabled -Value false -PSPath "IIS:\sites\$SiteName"
}


#---------------------------------------------------------[Request Cert]--------------------------------------------------------

Write-Host "Starting Challenge for " $dnsEntry "...";
Write-Host "Creating ACMEChallenge for IIS site " $SiteName "...";
Complete-ACMEChallenge $aliasCert -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = $SiteName }
$challenge = Submit-ACMEChallenge -Ref $aliasCert -Challenge http-01
While ($challenge.Status -eq "pending") {
  Start-Sleep -m 500 # wait half a second before trying
  Write-Host "Status is still 'pending', waiting for it to change..."
  $challenge = Update-ACMEIdentifier -Ref $aliasCert
}

if($challenge.Status -eq "invalid"){
    Write-Host "Status is invalid! Exiting."
    exit
}

Write-Host "Requesting Certificate from LetsEncrypt..."
New-ACMECertificate -Identifier $aliasCert -Alias $aliasCertExport -Generate
$certificateInfo = Submit-ACMECertificate -Ref $aliasCertExport

While([string]::IsNullOrEmpty($certificateInfo.IssuerSerialNumber)) {
 Start-Sleep -m 500 # wait half a second before trying
 Write-Host "IssuerSerialNumber is not set yet, waiting for it to be populated..."
 $certificateInfo = Update-ACMECertificate -Ref $aliasCertExport
}

cd $certStoragePath;

$certPath = $certStoragePath + $aliasCertExport + ".pfx";
#Save Certificate to disk
Write-Host "Writing Certificate to " $certPath
Get-ACMECertificate -Ref $aliasCertExport -ExportPkcs12 $certPath -CertificatePassword $certificiatePassword


#---------------------------------------------------------[Import Cert]--------------------------------------------------------

Write-Host 'Import pfx certificate from ' $certPath
$encryptedPassword = ConvertTo-SecureString $certificiatePassword –asplaintext –force 
Import-PfxCertificate -FilePath $certPath -CertStoreLocation "Cert:\$certRootStore\$certStore" -Password $encryptedPassword -Exportable

Write-Host "Recreating HTTPS Binding..."
$thumbprint = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath,$certificiatePassword)).Thumbprint
New-WebBinding -Name $SiteName -Protocol "https" -Port 443 -IPAddress * -HostHeader $dnsEntry -SSLFlags 1
$certificate = Get-ChildItem "Cert:\$certRootStore\$certStore" | Where-Object {$_.Thumbprint -eq $thumbprint}

Remove-Item -Path "IIS:\SslBindings\!443!$dnsEntry" -ErrorAction SilentlyContinue
New-Item -Path "IIS:\SslBindings\!443!$dnsEntry" -Value $certificate -SSLFlags 1

if($disableRewriteRule -eq 1){
    Write-Host "Enabling HTTPS Rewrite..."
    $rwStorN = '/system.webserver/rewrite/rules/rule[@name="{0}"]' -f $httpsRewriteRuleName
    set-webconfigurationproperty $rwStorN -Name enabled -Value true -PSPath "IIS:\sites\$SiteName"
}
#Reenable redirects if these were active.
if(redirectEnabled -eq "True"){
  Set-WebConfiguration system.webServer/httpRedirect "IIS:\sites\$SiteName" -Value @{enabled="true"} -ErrorAction SilentlyContinue
}

#---------------------------------------------------------[Cleaning Up]--------------------------------------------------------

if($deleteCertificateFileOnCompletion -eq 1){
 Write-Host 'Removing ' $certPath
 Remove-Item $certPath
}

#---------------------------------------------------------[Schedule Renewal]--------------------------------------------------------
if($scheduleRenewal -eq 1){
    Write-Host "Scheduling Renewal job..."
    $principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $arguments = "-File '{0}'" -f $scriptPath
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument $arguments
    $trigger =  New-ScheduledTaskTrigger -Daily -DaysInterval $scheduleDays -At $scheduleTime
    Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "IIS SSL Renewal for $dnsEntry" -Description "Renewal of SSL cert binding for Domain $dnsEntry on IIS site $SiteName" 

}
Write-Host 'Completed.'
