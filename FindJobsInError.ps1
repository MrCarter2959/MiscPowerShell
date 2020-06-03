######################################################
# Arrays
$taskArraysOK = New-Object System.Collections.Generic.List[object]

$taskArraysERR = New-Object System.Collections.Generic.List[object]

$taskArraysDIS = New-Object System.Collections.Generic.List[object]

$taskArraysDATE = New-Object System.Collections.Generic.List[object]

######################################################
# Set Task Folder Location
$taskFolder = "\PowerShell\*" #Enter the folder in which your tasks live

######################################################
# SMTP Options
$taskSettings = [PSCustomObject]@{
    From    = "email_address"
    To      = "email_address"
    Subject = "[ALERT] Scheduled Task in Error"
    Server  = "smtp_server"
}
######################################################
# Search Tasks
$findTasks = Get-ScheduledTask -TaskPath $taskFolder

ForEach ($task in $findTasks) {
        
        $taskExitCode = Get-ScheduledTask -TaskName $task.TaskName | Get-ScheduledTaskInfo | Select -ExpandProperty LastTaskResult

        $taskLastRunTime = Get-ScheduledTask -TaskName $task.TaskName | Get-ScheduledTaskInfo | Select -ExpandProperty LastRunTime

        [string]$taskLastRunDate = $taskLastRunTime.ToString("MM-dd-yyyy")

        $date = New-TimeSpan -Start $taskLastRunDate -End $taskDates

        $taskNextRunTime = Get-ScheduledTask -TaskName $task.TaskName | Get-ScheduledTaskInfo | Select -ExpandProperty NextRunTime

        $taskNumberOfMissedRuns = Get-ScheduledTask -TaskName $task.TaskName | Get-ScheduledTaskInfo | Select -ExpandProperty NumberOfMissedRuns

        $taskName = ($task.TaskName)
        
        $taskStatus = ($task.State)

        switch ($taskExitCode)
            {
                (0) {$taskExitDefination = "The operation completed successfully."}
                (1) {$taskExitDefination = "Incorrect function called or unknown function called."}
                (2) {$taskExitDefination = "File not found."}
                (10) {$taskExitDefination = "The environment is incorrect."}
                (267008) {$taskExitDefination = "Task is ready to run at its next scheduled time."}
                (267009) {$taskExitDefination = "Task is currently running."}
                (267010) {$taskExitDefination = "The task will not run at the scheduled times because it has been disabled."}
                (267011) {$taskExitDefination = "Task has not yet run."}
                (267012) {$taskExitDefination = "There are no more runs scheduled for this task."}
                (267013) {$taskExitDefination = "One or more of the properties that are needed to run this task on a schedule have not been set."}
                (267014) {$taskExitDefination = "The last run of the task was terminated by the user."}
                (267015) {$taskExitDefination = "Either the task has no triggers or the existing triggers are disabled or not set."}
                (2147750671) {$taskExitDefination = "Credentials became corrupted."}
                (2147750687) {$taskExitDefination = "An instance of this task is already running."}
                (2147943645) {$taskExitDefination = "The service is not available is (Run only when an user is logged on checked?)."}
                (3221225786) {$taskExitDefination = "The application terminated as a result of a CTRL+C."}
                (3228369022) {$taskExitDefination = "Unknown software exception."}
            
            }

            ######################################################
            # Tasks in Error State
            $taskInError = "1","10","267011","267012","267013","267014","267015","2147750671","2147750687","2147943645","3221225786","3228369022"
            ######################################################
            # Tasks in Success OR Running State
            $taskInStatus = "0","267008","267009"
            ######################################################
            # Tasks in Disabled State
            $taskInDisable = "267010"
            ######################################################
            # Check the Tasks We Found From Above
            #
            if ($date.Days -ge 1 -and $taskStatus -ne "Disabled")
                {
                    Write-Host "Task in Error: $taskName ; Days Since Last Run: $($date.Days) ; Exit Code: $TaskExitCode ; Reason: $taskExitDefination" -BackgroundColor "Yellow" -ForegroundColor "BlacK"
                    $taskArraysDATE.Add(
                        [PSCUSTOMOBJECT] @{Name=$taskName;TaskStatus=$taskStatus;DaysSinceLastRun=$($date.Days);TaskLastRunTime=$taskLastRunTime;NextRunTime=$taskNextRunTime;ExitCode=$taskExitCode;Definition=$taskExitDefination
                        }
                    )
                }
            Else
                {
                    #Write-Host "Statement was false"
                }
            ######################################################
            # For Tasks in Error State
            If ($TaskExitCode -in $taskInError)
                {
                    Write-Host "Task in Error: $taskName ; Exit Code: $TaskExitCode ; Reason: $taskExitDefination" -BackgroundColor "Yellow" -ForegroundColor "BlacK"
                    $taskArraysERR.Add(
                        [PSCUSTOMOBJECT] @{Name=$taskName;TaskStatus=$taskStatus;DaysSinceLastRun=$($date.Days);TaskLastRunTime=$taskLastRunTime;NextRunTime=$taskNextRunTime;ExitCode=$taskExitCode;Definition=$taskExitDefination
                        }
                    )
                }
            ######################################################
            # For Tasks in Success/Running State
            If ($taskExitCode -in $taskInStatus)
                {
                    Write-Host "Success On Task: $taskName ; Exit Code: $TaskExitCode ; Reason: $taskExitDefination" -BackgroundColor "Green" -ForegroundColor "Black"
                    $taskArraysOK.Add(
                        [PSCUSTOMOBJECT]@{Name=$taskName;TaskStatus=$taskStatus;DaysSinceLastRun=$($date.Days);TaskLastRunTime=$taskLastRunTime;NextRunTime=$taskNextRunTime;ExitCode=$taskExitCode;Definition=$taskExitDefination
                        }
                    )
                }
            ######################################################
            # For Tasks in Disabled State
            If ($taskExitCode -in $taskInDisable)
                {
                    Write-Host "Unable to run task: $taskName ; Exit Code: $TaskExitCode ; Reason: $taskExitDefination" -BackgroundColor "Red" -ForegroundColor "Black"
                    $taskArraysDIS.Add(
                        [PSCUSTOMOBJECT] @{Name=$taskName;TaskStatus=$taskStatus;DaysSinceLastRun=$($date.Days);TaskLastRunTime=$taskLastRunTime;NextRunTime=$taskNextRunTime;ExitCode=$taskExitCode;Definition=$taskExitDefination
                        }
                    )
                }
    }
######################################################
# SMTP Body
$taskSettingsBody = $taskArraysDATE | ConvertTo-Html -Head $taskHTML -Body "<H2>Tasks In Error</H2>" | Out-String
######################################################
# Send Alert
if ($taskArraysDATE.Count -ge 1) {
    Send-MailMessage -From $taskSettings.From -to $taskSettings.To -Subject $taskSettings.Subject -Body $taskSettingsBody -BodyAsHtml -SmtpServer $taskSettings.Server
}
