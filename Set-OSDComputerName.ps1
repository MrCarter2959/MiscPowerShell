[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")



#Set the computer name
Try {
    $TSProgressUI = New-Object -ComObject Microsoft.SMS.TSProgressUI
    $TSProgressUI.CloseProgressDialog()
    $TSProgressUI = $null
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $TSEnv.Value("SMSTSAssignUsersMode") = "Auto"
    If ($TSEnv.Value("OSDComputerName")) {
        $TSComputerName = $TSEnv.Value("OSDComputerName")
        }
        ElseIf ($TSEnv.Value("_SMSTSMACHINENAME")) {
            $TSComputerName = $TSEnv.Value("_SMSTSMACHINENAME")
            }
        Else {
            $TSComputerName = ""
            }
        if ($TSComputerName.Text.Length -gt 15) {
            $ErrorProvider.SetError($TSComputerName, "Computer name cannot exceed 15 characters")
            }
    }
    Catch {
        EXIT 1
        # Write-Error "$_"
        }

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider
 
$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(285,140)  
$Form.MinimumSize = New-Object System.Drawing.Size(285,140)
$Form.MaximumSize = New-Object System.Drawing.Size(285,140)
$Form.StartPosition = "CenterScreen"
$Form.SizeGripStyle = "Hide"
$Form.Text = "Enter Computer Name"
$Form.ControlBox = $false
$Form.TopMost = $true
 
$TBComputerName = New-Object System.Windows.Forms.TextBox
$TBComputerName.Location = New-Object System.Drawing.Size(25,30)
$TBComputerName.Size = New-Object System.Drawing.Size(215,50)
$TBComputerName.TabIndex = "1"
$TBComputerName.Text = $TSComputerName
 
$GBComputerName = New-Object System.Windows.Forms.GroupBox
$GBComputerName.Location = New-Object System.Drawing.Size(20,10)
$GBComputerName.Size = New-Object System.Drawing.Size(225,50)
$GBComputerName.Text = "Computer name:"
 
$ButtonOK = New-Object System.Windows.Forms.Button
$ButtonOK.Location = New-Object System.Drawing.Size(195,70)
$ButtonOK.Size = New-Object System.Drawing.Size(50,20)
$ButtonOK.Text = "OK"
$ButtonOK.TabIndex = "2"
$ButtonOK.Add_Click({Set-OSDComputerName})

$Form.KeyPreview = $True
$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){Set-OSDComputerName}})

Function Load-Form 
{
    $Form.Controls.Add($TBComputerName)
    $Form.Controls.Add($GBComputerName)
    $Form.Controls.Add($ButtonOK)
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}

Function Set-OSDComputerName 
{
    $ErrorProvider.Clear()
    if ($TBComputerName.Text.Length -eq 0) 
    {
        $ErrorProvider.SetError($GBComputerName, "Please enter a computer name.")
    }

    elseif ($TBComputerName.Text.Length -gt 15) 
    {
        $ErrorProvider.SetError($GBComputerName, "Computer name cannot be more than 15 characters.")
    }

    #Validation Rule for computer names.
    elseif ($TBComputerName.Text -match "^[-_]|[^a-zA-Z0-9-_]")
    {
        $ErrorProvider.SetError($GBComputerName, "Computer name invalid, please correct the computer name.")
    }

    else 
    {
        $OSDComputerName = $TBComputerName.Text.ToUpper()
        $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
        $TSEnv.Value("OSDComputerName") = "$($OSDComputerName)"
        $Form.Close()
    }
}

Load-Form
