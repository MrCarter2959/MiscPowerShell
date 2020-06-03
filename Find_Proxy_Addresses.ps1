$Users_To_Search = Import-CSV -Path "IMPORT_CSV" | % {

$Email_Alias_Looking_For = $_.EmailAddress_Norm #this is the alias we want to find which could potentially be unlisted in proxyaddresses, but also could be listed
$Primary_SMTP = $_.EmailAddress #this is the users SMTP:username@domain.org address
$Email_Alias_1 = $_.EmailAlias_1 #this is a alias from the old list
$Email_Alias_2 = $_.EmailAlias_2 #this is their smtp:username@domainname.mail.onmicrosoft.com alias from AD
$smtp_proxy = "smtp:$Email_Proxy"
$Alias_To_Find = "smtp:$Email_Alias_Looking_For"


$Find_My_User = Get-ADUser -Filter {(ProxyAddresses -eq $Primary_SMTP)} -Properties SamAccountName, ProxyAddresses, userPrincipalName | Select SamAccountName, ProxyAddresses, userPrincipalName


    Foreach ($user in $Find_My_User) {
        $samaccountname = ($user.samaccountname)
        $userprincipalname = ($user.userprincipalname)

        $Get_ProxyAddresses = Get-ADUser -Filter {(userPrincipalName -eq $userprincipalname)} -Properties ProxyAddresses | Select-Object @{ L = "ProxyAddresses"; E = {($_.ProxyAddresses | Where-Object {$_ -like "smtp:*" }) -join ';'} }
        $Seperate_Proxys = $Get_ProxyAddresses.ProxyAddresses.Split(';')
        $Seperate_Proxys
            foreach ($proxy in $Seperate_Proxys) {
            
                if ($proxy -cmatch $Email_Alias_1)
                    {Write-Host "Proxy : $Proxy matches Email Alias : $Email_Alias_1"-BackgroundColor "Gray" -ForegroundColor "Green"}
                if ($proxy -cmatch $Primary_SMTP)
                    {Write-Host "Proxy : $Proxy matches Email Alias : $Primary_SMTP" -BackgroundColor "Gray" -ForegroundColor "Black"}
                if ($proxy -cmatch $Email_Alias_2)
                    {Write-Host "Proxy : $Proxy matches Email Alias : $Email_Alias_2" -BackgroundColor "Gray" -ForegroundColor "Red"}
                if ($proxy -notmatch $Alias_To_Find)
                    {Write-Host "Proxy : $Proxy does not match Email Alias : $Alias_To_Find" -BackgroundColor "Gray" -ForegroundColor "Cyan"}
                #Else {
                #    Write-Host "Proxy : $Proxy does not match Email Alias : $Alias_To_Find" -BackgroundColor "Gray" -ForegroundColor "Magenta"}
            }
    }
}
