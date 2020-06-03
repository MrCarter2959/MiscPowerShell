Function Get-WordDocument()
{
    param (
        [string]$FullPath,
        [bool]$Visible = $False
    )
$openWord = New-Object -ComObject Word.Application
    If (!$openWord)
        {
            Write-Error -Message "Unable to open Word. Please check install."
        }
    Else
        {
            Write-Verbose "Word Installed"
        }

    Switch ($Visible)
        {
            $true { $openWord.Visible = $true}
            Default { $openWord.Visible = $false}
        }
$openWord.Visible = $Visible
Return $openWord.Documents.Add($FullPath)
Return $openWord
}

Function Close-WordDoucment ()
{
    $documents.close()
}

Function Quit-Word ()
{
    $openword.quit()
}

Function Save-WordDocument ()
{
      param (
        [string]$FilePath,
        [string]$FileName
    )
    $documents.SaveAs("$FilePath\$FileName.docx")

}

Function Replace-WordTemplate ()
{
    param (
        $WordDocument,
        [string]$FindText,
        [string]$ReplaceText
    )


    $FindReplace=$WordDocument
    $matchCase = $false;
    $matchWholeWord = $true;
    $matchWildCards = $false;
    $matchSoundsLike = $false;
    $matchAllWordForms = $false;
    $forward = $true;
    $format = $false;
    $matchKashida = $false;
    $matchDiacritics = $false;
    $matchAlefHamza = $false;
    $matchControl = $false;
    $read_only = $false;
    $visible = $true;
    $replace = 2;
    $wrap = 1;

$FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $ReplaceText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl) | Out-Null
}

Function Convert-WordToPDF ()
{
    param (
        [string]$FullPath,
        [bool]$Visible = $False
    )
$openWord = New-Object -ComObject Word.Application
    If (!$openWord)
        {
            Write-Error -Message "Unable to open Word. Please check install."
        }
    Else
        {
            Write-Verbose "Word Installed"
        }

    Switch ($Visible)
        {
            $true { $openWord.Visible = $true}
            Default { $openWord.Visible = $false}
        }    
    $openWord.Visible = $Visible
    $Doc = $openWord.Documents.Open($FullPath)
    $Name=($Doc.FullName).Replace('docx','pdf')
    $Doc.SaveAs($Name, 17)
    Remove-Item -Path $FullPath -Verbose
}

#Functions Require and Depend on $documents being = to Get-WordDocument for rest of functions to work.
$documents = Get-WordDocument -Visible $false -FullPath "C:\TEST-StudentTemplate.docx"

Replace-WordTemplate -WordDocument $documents.ActiveWindow.Selection.Find -FindText '<STUDENTNAME>' -ReplaceText 'password'
Replace-WordTemplate -WordDocument $documents.ActiveWindow.Selection.Find -FindText '<STUDENTUSERNAME>' -ReplaceText 'username'
Replace-WordTemplate -WordDocument $documents.ActiveWindow.Selection.Find -FindText '<STUDENTPASSWORD>' -ReplaceText 'password'

#Save the Document as DOCX, saving directly as PDF seems to cause issues
Save-WordDocument -FilePath "C:\" -FileName "username"
Close-WordDoucment
Quit-Word

#Testing way to force close word at the end of script
#Get-Process -Name WINWORD | Stop-Process

#Convert DOCX to PDF Format
Convert-WordToPDF -FullPath "C:\username.docx" -Visible $false
Close-WordDoucment
Quit-Word
