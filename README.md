# iis-letsencrypt-ssl-automation
SSL Automation for IIS Certificates via powershell script.

## How to
Configure the needed parameters directly within the [script](https://github.com/teriansilva/iis-letsencrypt-ssl-automation/blob/master/ssl-automation-script.ps1) and run it in powershell. It will automate your complete LetsEncrypt ssl renewal cycles.

## Requirements
You need to have ACMESharp Modules installed in Powershell. The [script](https://github.com/teriansilva/iis-letsencrypt-ssl-automation/blob/master/ssl-automation-script.ps1) will attempt to do this for you, if these are not installed but you might want to clarify before running the [script](https://github.com/teriansilva/iis-letsencrypt-ssl-automation/blob/master/ssl-automation-script.ps1).
Go here: https://github.com/ebekker/ACMESharp

## Other
I also added an example for a http to https rewrite rule. You can put it in the web.config to have your site automatically called via https. Make sure you have the URL Rewrite Module installed in your IIS!
