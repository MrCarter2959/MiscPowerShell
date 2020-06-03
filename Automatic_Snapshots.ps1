Import-Module -Name VMware.PowerCLI
###########################################################################################################################################################################################################
#
#Date Variables
$AutoMox_UniqueDate = Get-Date -Format "MM:dd:yyyy:HH"
$AutoMox_UniqueDate2 = Get-Date -Format "MM_dd_yyyy_HH"
$AutoMox_UniqueDate3 = Get-Date -Format "MM-dd-yyyy_HH"
$AutoMox_UniqueDate4 = Get-Date -Format "MM-dd-yyyy"
$AutoMox_CurrentDate = Get-Date -Format FileDateUniversal
#
###########################################################################################################################################################################################################
#
#SMTP Information
$AutoMox_SMTP = "smtp_server"
$AutoMox_To = "to_address"
$AutoMox_From = "from_address"
$AutoMox_Subject = "AutoMox Patch Snapshots on $AutoMox_UniqueDate"
#$AutoMox_CC = "email_address","email_address"
#
#
###########################################################################################################################################################################################################
#
#Log Location
$AutoMox_Log = "\\fileserver\path\PowerShell_Script_Log\AutoMoxSnapshots\$AutoMox_CurrentDate\"
$AutoMox_Output = "\\fileserver\path\PowerShell_Script_Log\AutoMoxSnapshots\$AutoMox_CurrentDate\AutoMox_Snapshot_Log_$Automox_UniqueDate2.txt"
#Create New Daily Folder
if (!(Test-Path -Path $AutoMox_Log )) {
    New-Item -ItemType Directory -Path $AutoMox_Log
        Write-Output "-------------------------------$AutoMox_UniqueDate-------------------------------" | Out-File $AutoMox_Output -Append
        Write-Host "New Folder Created For $AutoMox_Log" -BackgroundColor "Black" -ForegroundColor "Yellow"
        Write-Output "New Folder Created For $AutoMox_Log....Line 14" | Out-File $AutoMox_Output -Append
        }
#
$AutoMox_LogCSV = "\\fileserver\path\PowerShell_Script_Log\AutoMoxSnapshots\$AutoMox_CurrentDate\AutoMox_Snapshot_Log_$Automox_UniqueDate2.csv"
$AutoMox_LogHeader = "ServerName,SnapshotTaken,VMHost,ProvisonedSpace,UsedSpace,FreeSpace,DataStore,VMToolsVersion,VMToolsStatus,HardwareVersion,IPAddress,Date"
$AutoMox_LogHeader | Out-File $AutoMox_LogCSV -Encoding ASCII
#
#Email Attachment
$AutoMox_EmailAttachment = "\\fileserver\path\PowerShell_Script_Log\AutoMoxSnapshots\$AutoMox_CurrentDate\AutoMox_Snapshot_Log_$Automox_UniqueDate2.csv"
$AutoMox_EmailAttachments = @()
$AutoMox_EmailAttachments += $AutoMox_EmailAttachment
$AutoMox_EmailAttachments += $AutoMox_Output
#
###########################################################################################################################################################################################################
#
#vCenter Credentails
#
$AutoMox_VCenter = "vcenter-url.domain.org"
$AutoMox_Username = "username"
$AutoMox_Password = "password"
$AutoMox_Password_Secured = (ConvertTo-SecureString -AsPlainText $AutoMox_Password -Force)
$AutoMox_RemotePS_Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AutoMox_Username,$AutoMox_Password_Secured
#
###########################################################################################################################################################################################################
#
#Connect to vCenter
#Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false
#Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false
connect-viserver $AutoMox_VCenter -user $AutoMox_Username -password $AutoMox_Password
Write-Output "Connected to vCenter at https://$AutoMox_VCenter with $AutoMox_Username....Line 54" | Out-File $AutoMox_Output -Append
#
###########################################################################################################################################################################################################
#
#Find VM's Gather Information into Variables, Assign Snapshot Name
#
###########################################################################################################################################################################################################
$AutoMox_ADGroup = "AutoMox_Computers_Testing" # AD Group Name that contains computers that need a snapshot :)
#_Testing #--Uncomment and move _Testing to above variable to grab testing computers
$AutoMox_DomainController = "domain_controller"
$AutoMox_Comp = Get-ADGroupMember -Server $AutoMox_DomainController -Identity $AutoMox_ADGroup
$AutoMox_Name = ($AutoMox_Comp.Name)
$AutoMox_SnapShotName = "AutoMox-Updates-$AutoMox_UniqueDate4"
$AutoMox_SnapShotDescription = "Restore Point before Updates and Patches Installed by AutoMox on $AutoMox_UniqueDate. See AutoMox For List of Patches Applied. Restore To Here If Needed."
#
Write-Output "Starting WMI Query to Find Name and IP address......Line 51" | Out-File $AutoMox_Output -Append
Write-Output "---------------------------------------------------------------------------------" | Out-File $AutoMox_Output -Append
Foreach ($comp in $AutoMox_Comp ){
#
#$AutoMox_IP = Get-WmiObject -Class win32_networkadapterconfiguration -ComputerName ($comp.Name) | where { $_.ipaddress -like "10.*" } | select -ExpandProperty ipaddress | select -First 1
$AutoMox_CompName = ($comp.name)

$AutoMox_FQDN = Get-ADComputer -Server $AutoMox_DomainController -Filter {(Name -Like $AutoMox_CompName)} -Properties * | Select DNSHostName
$AutoMox_FQDN1 = ($AutoMox_FQDN.DNSHostName)

#Write-Output "WMI_Query to Find IP Address For $AutoMox_CompName...IP Address is $AutoMox_IP" | Out-File $AutoMox_Output -Append
#
$AutoMox_GetVM = Get-VM | Where-Object {$_.ExtensionData.Guest.HostName -like "*$($AutoMox_FQDN1)*"} | Select VMHost, ProvisionedSpaceGB, UsedSpaceGB, @{N=”Datastore”;E={[string]::Join(‘,’,(Get-Datastore -Id $_.DatastoreIdList | Select -ExpandProperty Name))}}, @{N="ToolsStatus";E={$_.Guest.Extensiondata.ToolsVersionStatus}}, HardwareVersion, Name
$AutoMox_GetVMTools = Get-VM | Where-Object {$_.ExtensionData.Guest.HostName -like "*$($AutoMox_FQDN1)*"} | Get-VMGuest | Select ToolsVersion , @{N="ToolsStatus";E={$_.Guest.Extensiondata.ToolsVersionStatus}}, IPAddress
$AutoMox_VMHost = ($AutoMox_GetVM.VMHost)
$AutoMox_ProvSpace = ($AutoMox_GetVM.ProvisionedSpaceGB)
$AutoMox_ProvSpace1 = [math]::Round($AutoMox_ProvSpace,3)
$AutoMox_ProvSpace2 = ("$AutoMox_ProvSpace1 GB")
$AutoMox_UsedSpace = ($AutoMox_GetVM.UsedSpaceGB)
$AutoMox_UsedSpace1 = [math]::Round($AutoMox_UsedSpace,3)
$AutoMox_UsedSpace2 = ("$AutoMox_UsedSpace1 GB")
$AutoMox_FreeSpace = ($AutoMox_ProvSpace1 - $AutoMox_UsedSpace1)
$AutoMox_FreeSpace1 = ("$AutoMox_FreeSpace GB")
$AutoMox_Datastore = ($AutoMox_GetVM.Datastore)
$AutoMox_Datastore1 = ($AutoMox_Datastore -replace "," , ":")
$AutoMox_Tools = ($AutoMox_GetVMTools.ToolsVersion)
$AutoMox_Hardware = ($AutoMox_GetVM.HardwareVersion)
$AutoMox_ToolsStatus = ($AutoMox_GetVM.ToolsStatus)
$AutoMox_VMName = ($AutoMox_GetVM.Name)
$AutoMox_IP = ($AutoMox_GetVMTools.IPAddress)
$AutoMox_GetVMIP1 = (Get-VM -Name $AutoMox_VMName).Guest.IPAddress[0]
$AutoMox_GetVMIP2 = (Get-VM -Name $AutoMox_VMName).Guest.IPAddress[1]
$AutoMox_GetVMIP3 = (Get-VM -Name $AutoMox_VMName).Guest.IPAddress[2]
$AutoMox_GetVMIP4 = (Get-VM -Name $AutoMox_VMName).Guest.IPAddress[3]
Write-Host "$AutoMox_VMName IP Addresses = $AutoMox_GetVMIP1 : $AutoMox_GetVMIP2 : $AutoMox_GetVMIP3"
Write-Host "Basic Info About $AutoMox_CompName..Host:$AutoMox_VMHost..ProvSpace:$AutoMox_ProvSpace2..UsedSpace:$AutoMox_UsedSpace2..FreeSpace:$AutoMox_FreeSpace1..Datastore:$AutoMox_Datastore1..VmToolsVersion:$AutoMox_Tools..VmToolsStatus:$AutoMox_ToolsStatus..HardwareVersion:$AutoMox_Hardware" -BackgroundColor "White" -ForegroundColor "Red"
Write-Output "Basic Info About $AutoMox_CompName..Host:$AutoMox_VMHost..ProvSpace:$AutoMox_ProvSpace2..UsedSpace:$AutoMox_UsedSpace2..FreeSpace:$AutoMox_FreeSpace1..Datastore:$AutoMox_Datastore1..VmToolsVersion:$AutoMox_Tools..VmToolsStatus:$AutoMox_ToolsStatus..HardwareVersion:$AutoMox_Hardware" | Out-File $AutoMox_Output -Append
###########################################################################################################################################################################################################
#
#Find Old Snapshop First and Delete it
#
###########################################################################################################################################################################################################
$AutoMox_OldSnapshot = (get-date).AddDays(-7).ToString("MM-dd-yyyy")
$AutoMox_OldSnapshotName = "AutoMox-Updates-$AutoMox_OldSnapshot"
$AutoMox_Blank = " "
Write-Host $AutoMox_OldSnapshotName #-Uncomment to See the Date the -1 is pulling
#

    Foreach ($snapshot in Get-VM -Name $AutoMox_VMName | Get-Snapshot -name $AutoMox_OldSnapshotName -ErrorVariable $err -ErrorAction SilentlyContinue ){
    #Write-Host $snapshot.Name
    $AutoMox_Snapname3 = ($snapshot.name)
    Write-Host $AutoMox_Snapname3
    #Write-Host "Getting Snapshot For $AutoMox_VMName. Looking For Snapshot $AutoMox_OldSnapshotName" -BackgroundColor "Yellow" -ForegroundColor "Black"
    #Write-Host ($AutoMox_GetSnapshot).name[0]
    #Write-Host ($AutoMox_GetSnapshot).name[1]
    if($AutoMox_Snapname3 -ceq $AutoMox_OldSnapshotName)
    {
        Write-Host $AutoMox_VMName "Snapshot Found for $AutoMox_Snapname3 Will Remove Snapshot..Line 133" -BackgroundColor "Cyan" -ForegroundColor "Black"
        $AutoMox_RemoveSnap = Get-VM -Name $AutoMox_VMName | Get-Snapshot -Name $AutoMox_Snapname3 | Remove-Snapshot -Confirm:$false
    }
    else
    {
        Write-Host "$AutoMox_VMName Has no snapshot with name of $AutoMox_OldSnapshotName..Line 138" -BackgroundColor "White" -ForegroundColor "DarkCyan"
    }
    #
    If ($AutoMox_SnapName3 -match $AutoMox_SnapShotName)
    {
        Write-Host "Snapshot Already Exists Under $AutoMox_SnapName3, Won't Take One... Line 139"
    }
    else
    {
        Write-Host "Snapshot Not Found For $AutoMox_SnapShotName. Will Take Snapshot Under $AutoMox_SnapShotName..Line 143" -BackgroundColor "Black" -ForegroundColor "Magenta"
        Get-VM -Name $AutoMox_VMName | New-Snapshot -Name $AutoMox_SnapShotName -Description $AutoMox_SnapShotDescription -Memory
        Write-Host "Snapshot taken for $AutoMox_VMname with Description: $AutoMox_SnapShotDescription and options -Memory..Line 145" -BackgroundColor "White" -ForegroundColor "DarkCyan"
    }
    If ($AutoMox_SnapName3 -ceq $AutoMox_SnapShotName)
    {
    Write-Host "Snapshot Already Exists Under $AutoMox_SnapshotName, Won't Take One..." -BackgroundColor "Gray" -ForegroundColor "Cyan"
    }
    Break
}
#For Machines that have no snapshots to compare against
    $AutoMox_Snap4 =  Get-VM -Name $AutoMox_VMName | Get-Snapshot -Name $AutoMox_SnapshotName -ErrorAction SilentlyContinue | Select -ExpandProperty Name #--Compares Todays Snapshot against Machines without Snapshots
    #foreach ($snapshot2 in Get-VM -Name "Papercut Web Print" | Get-Snapshot -Name $AutoMox_OldSnapshotName -ErrorVariable $err1 -ErrorAction SilentlyContinue | Select Name)
    $AutoMox_Snap5 = ($AutoMox_Snap4 -eq $AutoMox_SnapshotName)
    if ($AutoMox_Snap5 -eq $True)
    {
        Write-Host "$AutoMox_VMName Has Snapshot For $AutoMox_SnapshotName for $AutoMox_SnapshotName.....Line 152" -BackgroundColor "White" -ForegroundColor "DarkBlue"
    }
    If ($AutoMox_Snap5 -eq $False)
    {
        Write-Host "No Snapshots available for Machine $AutoMox_VMName, Taking $AutoMox_SnapShotName.....Line 152" -BackgroundColor "White" -ForegroundColor "DarkBlue"
        New-Snapshot -VM $AutoMox_VMName -Name $AutoMox_SnapShotName -Description $AutoMox_SnapShotDescription -Memory
        Write-Host "Snapshot taken for $AutoMox_VMname with Description: $AutoMox_SnapShotDescription and options -Memory..Line 166" -BackgroundColor "White" -ForegroundColor "DarkCyan"
    }
#
Write-Output "---------------------------------------------------------------------------------" | Out-File $AutoMox_Output -Append
$AutoMox_CSVAdd = $AutoMox_FQDN1 + "," + $AutoMox_SnapShotName + "," + $AutoMox_VMHost + "," + $AutoMox_ProvSpace2 + "," + $AutoMox_UsedSpace2 + "," + $AutoMox_FreeSpace1 + "," + $AutoMox_Datastore1 + "," + $AutoMox_Tools + "," + $AutoMox_ToolsStatus + "," + $AutoMox_Hardware + "," + $AutoMox_IP
$AutoMox_CSVAdd | Out-File $AutoMox_LogCSV -Append -Encoding ASCII
#

###########################################################################################################################################################################################################
#Teams Notifications
$Patch_vCenter_uri = 'Teams_WebHook_URI'

# Time
$Patch_vCenter_Time = get-date -format "HH:mm-MM/dd/yyyy"

# Script Name
$Patch_vCenter_Script_Name = "script_name"

# Script Name
$Patch_vCenter_RealName = "script_name without .ps1"

#Snapshot Name
$Patch_vCenter_Snapshot_Name = "BeforePatchWindow_$Patch_CurDate"

# these values would be retrieved from or set by an application

$Patch_vCenter_Teams_Body = ConvertTo-Json -Depth 4 @{
  title    = "$Patch_vCenter_Script_Name Completed Successfully"
  text	 = " "
  sections = @(
    @{
      activityTitle    = "$AutoMox_VMName Snapshots Taken"
      activitySubtitle = 'Patch Snapshots'
      #activityText	 = ' '
      activityImage    = 'image_URL' # this value would be a path to a nice image you would like to display in notifications
    },
    @{
      title = '<h2 style=color:blue;>vCenter Snapshot Details'
      facts = @(
        @{
          name = 'vCenter Username'
          value = $AutoMox_Username
          },
        @{
          name = 'Provisioned Space'
          value = $AutoMox_ProvSpace2
          },
        @{
          name = 'Used Space'
          value = $AutoMox_UsedSpace2
          },
        @{
          name = 'Free Space'
          value = $AutoMox_FreeSpace1
          },
        @{
          name = 'VM Datastores'
          value = $AutoMox_Datastore1
          },
        @{
          name = 'VM Tools'
          value = $AutoMox_Tools
          },
        @{
          name = 'VM Tools Status'
          value = $AutoMox_ToolsStatus
          },
        @{
          name = 'VM Hardware Version'
          value = $AutoMox_Hardware
         },
        @{
          name = 'VM IP #1'
          value = $AutoMox_GetVMIP1
         },
        @{
          name = 'VM IP #2'
          value = $AutoMox_GetVMIP2
          },
        @{
          name = 'VM IP #3'
          value = $AutoMox_GetVMIP3
         },
        @{
          name = 'VM IP #4'
          value = $AutoMox_GetVMIP4
         },
        @{
          name = 'Snapshot Taken'
          value = "$AutoMox_SnapShotName"
          },
        @{
          name = 'Snapshot Removed'
          value = $AutoMox_OldSnapshotName
          }
      )
    }
  )
}
Invoke-RestMethod -uri $Patch_vCenter_uri -Method Post -body $Patch_vCenter_Teams_Body -ContentType 'application/json'
#
}

###########################################################################################################################################################################################################
#
#Disconnect from vCenter
Disconnect-VIServer $AutoMox_VCenter -Confirm:$False
Write-Host "Disconnected from $AutoMox_VCenter with Credentials $AutoMox_Username" -BackgroundColor "Black" -ForegroundColor "White"
Write-Output "Disconnected From $AutoMox_VCenter.....Line 231" | Out-File $AutoMox_Output -Append
Write-Output "---------------------------------------------------------------------------------" | Out-File $AutoMox_Output -Append
#
###########################################################################################################################################################################################################
