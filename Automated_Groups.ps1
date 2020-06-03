# Import Modules. Only needs to be done once in the script
Import-Module ActiveDirectory

$employeeData = Import-CSV -Path "UNC_PATH"
######Uncomment for script testing######## "C:\PowerShell_Test\Automated_Group_Test.csv"

# Define some variables needed for overall code
$employeeDC = "DOMAIN_CONTROLLER"
$employeeArray = @()
$emailMatch = "*EMAIL_DOMAIN*" # Shouldn't use $Match as it's a reserved variable
$nonEmployeeGroups = @(
    "Student_License",
    "Alumni_License"
)
$nonEmployeeGroupMembers = @()

# Config for matching the employee type to the group name
$config = @{
    "Administrative"   = @{ GroupName = "Administration"; Members = @() }
    "Support Staff"    = @{ GroupName = "Support Staff"; Members = @() }
    "Faculty"          = @{ GroupName = "Faculty"; Members = @() }
    "Full-Time"        = @{ GroupName = "FT-Employees"; Members = @() }
    "Part-Time"        = @{ GroupName = "PT-Employees"; Members = @() }
    "Occasional-Staff" = @{ GroupName = "Occasional-Staff"; Members = @() }
    "Retired"          = @{ GroupName = "Retired"; Members = @() }
}

# Loop through the config and get all the group members now to save on execution time later. No point in getting the group members for each user
foreach ($employeeType in $config.Keys) {
    $config[$employeeType]['Members'] += Get-ADGroupmember -Identity $config[$employeeType]['GroupName'] -Server $employeeDC -Recursive | Select-Object -ExpandProperty samaccountname
}

# Get group members from AD for non-employees
foreach ($group in $nonEmployeeGroups) {
    $nonEmployeeGroupMembers += Get-ADGroupMember -Identity $group -Server $employeeDC -Recursive | Select-Object -ExpandProperty samaccountname
}

# Loop through employee data
foreach ($employee in $employeeData) {
    # Add some data onto the object for easier processing
    $employee | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "$($employee.LAST_NAME), $($employee.FIRST_NAME)"
    $employee | Add-Member -MemberType NoteProperty -Name "Name" -Value "$($employee.FIRST_NAME) $($employee.LAST_NAME)"

    # Check if the employee e-mail is not blank
    if ($employee.WORK_EMAIL -ne "") {
        # Find the employee in AD
        $employeeADAccount = Get-ADUser -Filter { (mail -eq $($employee.WORK_EMAIL)) -and (Enabled -eq $true) } -Server $employeeDC -Properties samaccountname, distinguishedname | Select-Object samaccountname, distinguishedname
    } elseif ($employee.WORK_EMAIL -eq "") {
        $employeeADAccount = Get-ADUser -Filter { (DisplayName -eq $employee.DisplayName) -and (Enabled -eq $true) } -Server $employeeDC -Properties samaccountname, distinguishedname, mail | Select-Object samaccountname, distinguishedname, mail
    }

    # Check to make sure the employee is not a member of the non employee groups and the e-mail matches
    if ($nonEmployeeGroupMembers -notcontains $employeeADAccount.samaccountname -and $employee.WORK_EMAIL -like $emailMatch) {
        # Check to make sure the employee type is in the config. There could be issues if we encounter an employee type we don't know about
        if ($config.Keys -contains $employee.EMPLOYEE_TYPE) {
            # Check to see if the employee is already a member of the group
            if ($config[$employee.EMPLOYEE_TYPE]['Members'] -notcontains $employeeADAccount.samaccountname) {
                # Output employee to add
                Write-Host "Add $($employee.WORK_EMAIL) to $($config['$employee.EMPLOYEE_TYPE']['GroupName'])" -Background "Green" -Foreground "Black"
                
                # Add employee to group
                Add-ADGroupMember -Identity $config['$employee.EMPLOYEE_TYPE']['GroupName'] -Members $employeeADAccount.distinguishedname -Verbose
                
                # Add employee to employee array
                $employeeArray += [PSCustomObject]@{
                    Username  = $employee.WORK_EMAIL
                    GroupName = $config['$employee.EMPLOYEE_TYPE']['GroupName']
                }
            } else {
                # Write that the employee is a member
                Write-Host "$($employee.WORK_EMAIL) is a member of Group: $($config['$employee.EMPLOYEE_TYPE']['GroupName'])" -Background "White" -Foreground "Black"
            }
        }
    }
}

###########################################
#             HTML Varables               #
###########################################
$employeeHtml = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}
</style>
"@
$employeeEmailBody = $employeeArray | ConvertTo-Html -Head $employeeHtml -Body "<H2>New Group Members</H2>" | Out-String
###########################################
#            SMTP Variables               #
###########################################
$smtpSettings = [PSCustomObject]@{
    From    = "FROM_ADDRESS"
    To      = "TO_ADDRESS"
    Subject = "AD Group Membership"
    Server  = "SMTP_SERVER"
}
##########################################
#              Send the Mail             #
##########################################
if ($employeeArray.Count -ge 1) {
    #Send-MailMessage -From $smtpSettings.From -to $smtpSettings.To -Subject $smtpSettings.Subject -Body $employeeEmailBody -BodyAsHtml -SmtpServer $smtpSettings.Server

    # Vars for REST
    $employeeUri = 'TEAMS_URI'
    $employeeTime = Get-Date -format "MM-dd-yyyy HH:mm"
    $employeeScriptName = "Automated_Groups.ps1"
    $employeeRealName = "Automated_Groups"

    # these values would be retrieved from or set by an application

    $employeeTeamsBody = ConvertTo-Json -Depth 4 @{
        title    = "$employeeScriptName Completed Successfully"
        text     = " "
        sections = @(
            @{
                activityTitle    = 'Automated_Groups Completed'
                activitySubtitle = 'Automated Group Membership Updates'
                #activityText	 = ' '
                activityImage    = 'PNG IMAGE' # this value would be a path to a nice image you would like to display in notifications
            },
            @{
                title = '<h2 style=color:blue;>Script Details'
                facts = @(
                    @{
                        name  = 'ScriptName'
                        value = $employeeScriptName
                    },
                    @{
                        name  = 'Finished'
                        value = $employeeTime
                    },
                    @{
                        name  = 'Group Additionals'
                        value = $employeeEmailBody
                    }
                )
            }
        )
    }
    # Send REST call to Teams
    Invoke-RestMethod -uri $employeeUri -Method Post -body $employeeTeamsBody -ContentType 'application/json'
}
