###############################
####    Import Function    ####
###############################
Function Connect-ExchangeOnPrem {
    
    #Exchange Username
    $Exchange_Username = "username"
    
    #Exchange Password
    $Exchange_Password = "password"
    
    #Exchange Password Secured
    $Exchange_Password1 = ConvertTo-SecureString $Exchange_Password -AsPlainText -Force
    
    #Exchange Admin Credentials
    $Exchange_adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $Exchange_Username,$Exchange_Password1

    #User Login
    $Login_Credentials = $Exchange_adminCredential

    #Exchange URL
    $Exchange_URI = "http://exchange_URL/PowerShell/"

    #Config Name
    $Configuration_Name = "Microsoft.Exchange"

    #Authentication
    $Exchange_Auth = "Kerberos"

    #Start The Session
    $Exchange_Session = New-PSSession -ConfigurationName $Configuration_Name -ConnectionUri $Exchange_URI -Authentication $Exchange_Auth -Credential $Login_Credentials

    #Import The Session
    Import-PSSession $Exchange_Session -DisableNameChecking -AllowClobber
    }
###############################
####    Import Modules     ####
###############################
Import-Module ActiveDirectory
###############################
####    Starting Script   #####
###############################

Connect-ExchangeOnPrem

$Mail_Contact_Array = @()

$Mail_Contact_CSV = Import-CSV -Path "csv_path"


    Foreach ($Contact in $Mail_Contact_CSV) 
        
        {
            $Mail_Contact_F_Name = $Contact.FirstName

            $Mail_Contact_L_Name = $Contact.LastName

            $Mail_Contact_Email = $Contact.EmailAddress

            $Mail_Contact_Requester = $Contact.Requester

            $domainName = "@domain.org"

            $Mail_Contact_Group_Name = $Contact.GroupsToAdd.split(",")

            $Mail_Contact_Name = $Mail_Contact_F_Name +" "+ $Mail_Contact_L_Name +" "+ $domainName

            $Mail_Contact_OU = 'OU=Some,OU=Path,DC=DOMAIN,DC=ORG'

            $Mail_Contact_Alias = $Mail_Contact_Email.Split("@")[0]

            $Mail_Contact_Check = Get-Recipient -Identity $Mail_Contact_Email

            If ($Mail_Contact_Check -ne "")

                {
                    New-MailContact -Name $Mail_Contact_Name -DisplayName $Mail_Contact_Name -FirstName $Mail_Contact_F_Name -LastName $Mail_Contact_L_Name -Alias $Mail_Contact_Alias -ExternalEmailAddress $Mail_Contact_Email -OrganizationalUnit $Mail_Contact_OU -Verbose

                    Get-MailContact -Identity $Mail_Contact_Name | Set-Contact -Company "Company" -Department "Department" -Title "job Title"

                    $Mail_Contact_Add = @"
                    EmailAddress,OU
                    "",""
                    $Mail_Contact_Email,$Mail_Contact_OU
"@ | ConvertFrom-CSV
            
                    $Mail_Contact_Array += $Mail_Contact_Add

                    Foreach ($group in $Mail_Contact_Group_Name)
                        {
                            Set-ADGroup -Identity $group -Add @{'member'="DN OF CONTACT"} -verbose
                        }
                    

                }
            Else 
                {
                    Write-Host "Mail Contact Already Created for $Mail_Contact_Email"

                    #check ad groups
                        foreach ($group in $Mail_Contact_Group_Name)
                            {

                                $compare = Get-ADGroupMember -Identity $group | Get-ADObject | Select mail
                                $findMissing = Compare-Object -ReferenceObject $compare -DifferenceObject $Mail_Contact_Email | Where $_.SideIndicator -eq "=>"
                                
                                $findDNMissing = Get-ADObject -filter "mail -eq "$findMissing.inputObject"" | Select -ExpandProperty DistinguishedName
                                Set-ADGroup -Identity $group -Add @{'member'="$findDNMissing"} -verbose
                            }

                  # OTher Items here, check DisplayName etc.. Create New IF Statement for each field you want checked similary to below
                  If ($Mail_Contact_Check.displayName -ne $Mail_Contact_Name)
                    {
                      Set-Contact -Identitiy "NAME/DN" -DisplayName $Mail_Contact_Name
                    }
                  Else
                    {
                      Write-Host "NAME/DN already has required displayName"
                    }
                }
        }
#Remove Exchange Connection
Get-PSSession | Remove-PSSession -Verbose
