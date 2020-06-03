

#New CSV with correct data and data is unique
$Aliases_New = Import-CSV -LiteralPath "IMPORT_PATH"


#Start the Loop
Foreach($Proxy_Aliases in $Aliases_New){
    
    #Assign SamAccountName from the csv to variable
    $Proxy_SamAccountName = $Proxy_Aliases.SamAccountName
    
    #Search using Get-ADUser
    $Proxy_CurrentUser = Get-ADUser -Filter {(SamAccountName -eq $Proxy_SamAccountName) -and (Enabled -eq $True)} -Properties ProxyAddresses, SamAccountName | Select ProxyAddresses, SamAccountName
    
    #Assign SamAccountName for $Proxy_CurrentUser
    $Proxy_CurrentUser_SAM = $Proxy_CurrentUser.SamAccountName
    
    #Assign the Proxys that could be missing and filter on where
    $MissingProxyAddresses =  $Proxy_Aliases.EmailAddress,$Proxy_Aliases.EmailAlias_1,$Proxy_Aliases.EmailAlias_2,$Proxy_Aliases.EmailAlias_2,$Proxy_Aliases.EmailAlias_3,$Proxy_Aliases.EmailAlias4 | where{$_ -NotIn $Proxy_CurrentUser.ProxyAddresses}
	
    #if $MissingProxys has 1 address do the following
    if ($MissingProxyAddresses -gt 1) {
        
        #Assign the Single address and split on comma
        $Single = ($MissingProxyAddresses -split ',')
        
        #assign to the custom object and export
        [PSCUSTOMOBJECT]@{
		    Username = $Proxy_CurrentUser_SAM
		    Alias1 = $Single[0]
            Alias2 = $Single[1]
            Alias3 = $Single[2]
            Alias4 = $Single[3]
            Alias5 = $Single[4]
	    } | Export-CSV -Path "EXPORT_PATH" -Append -NoTypeInformation
    }
}
