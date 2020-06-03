# Functions Used In Script
Function Connect-ExchangeOnPrem {
    
    #User Login
    $Login_Credentials = Get-Credential

    #Exchange URL
    $Exchange_URI = "http://EXCHANGE_URL/PowerShell/"

    #Config Name
    $Configuration_Name = "Microsoft.Exchange"

    #Authentication
    $Exchange_Auth = "Kerberos"

    #Start The Session
    $Exchange_Session = New-PSSession -ConfigurationName $Configuration_Name -ConnectionUri $Exchange_URI -Authentication $Exchange_Auth -Credential $Login_Credentials

    #Import The Session
    Import-PSSession $Exchange_Session -DisableNameChecking -AllowClobber
    }

function Get-SMTPAddress {
    param(
    [string]$DisplayName
    )
    $filter = "DisplayName -eq '$DisplayName'"
    (Get-ADObject -Filter $filter -Properties Mail | Select Mail)
}

# Set Array To Log To
$array = New-Object System.Collections.Generic.List[object]

# Connect to On Prem Exchange Environment
Connect-ExchangeOnPrem

# CSV To Import
$csvList = Import-CSV -Path "IMPORT_PATH"


forEach ($dlList in $csvList)
{
    $dLName = ($dlList.DistrubutionList.Split(","))

    Foreach ($distro in $dLName)
        {
            $array2 = New-Object System.Collections.Generic.List[object]

            $dlUser = ($dlList.Users)

            Write-Host "Finding Approved Senders For: $distro" -Background "Yellow" -ForegroundColor "Black"

            Start-Sleep -s 1

            $dList = Get-DistributionGroup -Identity $distro | Select AcceptMessagesOnlyFromSendersOrMembers

            ForEach ($user in $dList.AcceptMessagesOnlyFromSendersOrMembers) { 
                
                $mail = Get-SMTPAddress -DisplayName $user.split('/')[-1]

                $type = (Get-ADObject -Filter "mail -eq '$($mail.Mail)'" | Select ObjectClass)

                $GroupName = $Mail.Mail.Split("@")[0]

                Switch ($Type.ObjectClass)
                    {
                        "Group"    { $Explain = "Contains Multiple Users" }
                        "Contact"  { $Explain = "Mail Contact, Allows External Mail Users to Send to Email Lists" }
                        "User"     { $Explain = "Active Directory User, Normal User Account"}
                    }

                If ($type.ObjectClass -eq 'group')
                    {                        
                        $groupMembers = (Get-DistributionGroupMember -Identity $mail.Mail).Name
                        
                        Foreach ($member1 in $groupMembers) {

                            $membersFirst = $member1.Split(",")[1]
                            
                            If ($membersFirst) { $membersFirst = $membersFirst.Trim() }
                            
                            $membersLast = $member1.Split(",")[0]
                            
                            If ($membersLast) { $membersLast = $membersLast.Trim() }
                            # Formats FirstName LastName, instead of LastName, FirstName for easy reading
                            $DisplayName = ($membersFirst +" "+$membersLast).Trim()
                            #Write-Host $DisplayName
                            $array2.Add($DisplayName)
                        }
                        
                        If ($groupMembers)
                            {
                                $array.Add(
                                    [PSCUSTOMOBJECT]@{DistrubtionList=$distro;AllowedSenders=$($mail.Mail);Type=$($Type.ObjectClass);TypeExplaination=$Explain;"$GroupName Members"=($array2 | Out-String)}
                                )
                            }
                        }
                Else{
                    $array.Add(
                        [PSCUSTOMOBJECT]@{DistrubtionList=$distro;AllowedSenders=$($mail.Mail);Type=$($Type.ObjectClass);TypeExplaination=$Explain}
                    )
                }
                Write-Host "Found: $($mail.Mail)" -BackgroundColor "Cyan" -ForegroundColor "BlacK"
            }
        }
}


$HTML1 = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:white}
</style>
"@

$htmlBody1 = $array | ConvertTo-Html -Head $HTML1 -Body "<H2>Allowed Sender Reports</H2>" | Out-String

# Testing Sending Mail ALert, will format to use @splatting or array objects in end
Send-MailMessage -SmtpServer 'SMTP_SERVER' -To 'TO' -From 'FROM' -Subject "SUBJECT" -Body $htmlBody1 -BodyAsHtml

Get-PSSession | Remove-PSSession
