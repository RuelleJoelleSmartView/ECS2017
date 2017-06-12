# joelle.ruelle@smartview2.onmicrosoft.com

#Create authentification object for the user
$credential = Get-Credential -UserName 'joelle.ruelle@smartview2.onmicrosoft.com' -Message "Enter SPO credentials"

#Initializing a persistent connection to Exchange
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRed

#Import of the session in the current session
Connect-MsolService -Credential $credential

#Connect to Azure Active Directory with credentials 
Import-PSSession $Session -AllowClobber

#########################################################################################################################

#Get user by DisplayName
$Myuser = Get-MsolUser  | Where-Object {$_.DisplayName -eq "Joelle Ruelle"}

$DisplayName= “$($Myuser.DisplayName)”
Write-Host "DisplayName -> " $DisplayName
$Title = "$($Myuser.Title)"
Write-Host "Title -> " $Title
$MobilePhone = "$($Myuser.MobilePhone)"
Write-Host "MobilePhone -> " $MobilePhone
$UserPrincipalName = "$($Myuser.UserPrincipalName)"
Write-Host "UserPrincipalName -> " $UserPrincipalName

#Create HTML string signature format for selected user
$signHTML="<span style=`"font-family: calibri,sans-serif;`"><strong>" + $DisplayName + "</strong> - " + $Title
$signHTML+="<br />"
$signHTML+= "Phone: " + $MobilePhone 
$signHTML+="<br />"
$signHTML+= "Mail: " + $UserPrincipalName
$signHTML+="<br />"
$signHTML+="</span><br /><strong> Join us at the <a href='http://www.collaborationsummit.rocks/'> EUROPEAN COLLABORATION SUMMIT 2017 </a> - Zagreb, Croatia | May 29-31 2017"

#Assign signature to the selected user 
Set-MailboxMessageConfiguration –Identity $Myuser.UserPrincipalName -AutoAddSignature $True  -SignatureHtml   $signHTML


Write-Host "Déconnexion"
get-PSSession | remove-PSSession