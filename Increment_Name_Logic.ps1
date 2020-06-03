	########################################################################################################################################################################################
	#                                                                              Incrementing Name Logic                                                                                 #
	$StudentAccounts_Name1 = $StudentAccounts_Username # Enter Username variable here
	$StudentAccounts_DefaultName = $StudentAccounts_Name1 
	$StudentAccounts_Input_Exit = 0
	$StudentAccounts_Input_Count = 1
	########################################################################################################################################################################################
	#Generates Username if duplicates detected
	#######################################################################################################################################################################################
	# Change this line to fit what properites you need
	$StudentAccounts_CheckAD = Get-ADUser -Server $StudentDomainController -Filter {(SamAccountName -eq $StudentAccounts_Username) -and (EmployeeID -eq $StudentAccounts_StudentNumber)} -Properties SamAccountName, EmployeeID | Select SamAccountName, EmployeeID
	#
  #This Compares and checks AD from the line Above
	If ($StudentAccounts_CheckAD.SamAccountName -ne $StudentAccounts_Username -and $StudentAccounts_CheckAD.EmployeeID -ne $StudentAccounts_StudentNumber)
	{
	Do
	{
	    Try
	        {
	            #Attempts to get users info
	            $StudentAccounts_User = Get-ADUser -Server $StudentDomainController -Identity $StudentAccounts_Name1 -Properties SamAccountName | Select SamAccountName
	            Write-Host "Searching AD For $StudentAccounts_Name1" -ForegroundColor "Magenta"
	            #
	            #The User exists
	            #
	            $StudentAccounts_Name1 = $StudentAccounts_DefaultName + $StudentAccounts_Input_Count++
	            Write-Host $StudentAccounts_User.SamAccountName "Was a duplicate. AD Has a Account with:" $StudentAccounts_User.SamAccountName", New Username is: $StudentAccounts_Name1" -ForegroundColor "Green"
	            $StudentAccounts_Username = $StudentAccounts_Name1
	            If ($StudentAccounts_Count -gt 1) {$StudentAccounts_Input_Exit = 1}
	        }
	    Catch
	        {
	            $StudentAccounts_Input_Exit = 1
	            #Write-Host "Name Doesn't Exist in AD" -ForegroundColor "Green"
	            Break
	        }
	    } Until ($StudentAccounts_Input_Exit -eq 0)
	        Write-Host "$StudentAccounts_Name1 Wasn't a Duplicate" -ForegroundColor "Magenta"
	        Write-Host "$StudentAccounts_Name1 Needs a AD Account" -ForegroundColor "Yellow"
	}
