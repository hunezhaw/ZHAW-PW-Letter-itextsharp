Import-Module AGPLGeneratePWLetters
Import-Module AGPLLogging


Function Create-Form(){

	# define return object
	$retObj = New-Object Object
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'filter' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'validfrom' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'validto' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'sendto' -Value ''

	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = "Filter Entry Form"
	$objForm.Size = New-Object System.Drawing.Size(300,230) 
	$objForm.StartPosition = "CenterScreen"

	$objForm.KeyPreview = $True
	$objForm.Add_KeyDown(
		{
			if ($_.KeyCode -eq "Enter") 
			{
				$retObj.filter=$objTextFilter.Text;
				$retObj.validfrom=$objTextValidFrom.Text;
				$retObj.validto=$objTextValidTo.Text;
				$retObj.sendto=$objTextSendTo.Text;
				$objForm.Close()
			}
		})
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close()}})

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(75,150)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
	$OKButton.Add_Click({
							$retObj.filter=$objTextFilter.Text;
							$retObj.validfrom=$objTextValidFrom.Text;
							$retObj.validto=$objTextValidTo.Text;
							$retObj.sendto=$objTextSendTo.Text;
							$objForm.Close()
						})
	$objForm.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(150,150)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$objForm.Close()})
	$objForm.Controls.Add($CancelButton)

	$objLabelFilter = New-Object System.Windows.Forms.Label
	$objLabelFilter.Location = New-Object System.Drawing.Size(10,20) 
	$objLabelFilter.Size = New-Object System.Drawing.Size(60,20) 
	$objLabelFilter.Text = "Filter:"
	$objForm.Controls.Add($objLabelFilter) 

	$objTextFilter = New-Object System.Windows.Forms.TextBox 
	$objTextFilter.Location = New-Object System.Drawing.Size(80,18) 
	$objTextFilter.Size = New-Object System.Drawing.Size(190,20) 
	$objForm.Controls.Add($objTextFilter) 
	
	$objLabelValidFrom = New-Object System.Windows.Forms.Label
	$objLabelValidFrom.Location = New-Object System.Drawing.Size(10,50) 
	$objLabelValidFrom.Size = New-Object System.Drawing.Size(60,20) 
	$objLabelValidFrom.Text = "ValidFrom:"
	$objForm.Controls.Add($objLabelValidFrom)
	
	$objTextValidFrom = New-Object System.Windows.Forms.TextBox 
	$objTextValidFrom.Location = New-Object System.Drawing.Size(80,48) 
	$objTextValidFrom.Size = New-Object System.Drawing.Size(190,20) 
	$objForm.Controls.Add($objTextValidFrom) 
	
	$objLabelValidTo = New-Object System.Windows.Forms.Label
	$objLabelValidTo.Location = New-Object System.Drawing.Size(10,80) 
	$objLabelValidTo.Size = New-Object System.Drawing.Size(60,20) 
	$objLabelValidTo.Text = "ValidTo:"
	$objForm.Controls.Add($objLabelValidTo)	

	$objTextValidTo = New-Object System.Windows.Forms.TextBox 
	$objTextValidTo.Location = New-Object System.Drawing.Size(80,78) 
	$objTextValidTo.Size = New-Object System.Drawing.Size(190,20) 
	$objForm.Controls.Add($objTextValidTo) 
	
	$objLabelSendTo = New-Object System.Windows.Forms.Label
	$objLabelSendTo.Location = New-Object System.Drawing.Size(10,110) 
	$objLabelSendTo.Size = New-Object System.Drawing.Size(60,20) 
	$objLabelSendTo.Text = "SendTo:"
	$objForm.Controls.Add($objLabelSendTo)	

	$objTextSendTo = New-Object System.Windows.Forms.TextBox 
	$objTextSendTo.Location = New-Object System.Drawing.Size(80,108) 
	$objTextSendTo.Size = New-Object System.Drawing.Size(190,20) 
	$objForm.Controls.Add($objTextSendTo) 	

	$objForm.Topmost = $True

	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
	
	return $retObj
}

# get the parameters
$dialoginput = Create-Form

# assign and check values
$filter = $dialoginput.filter
if ($dialoginput.validfrom) {[DateTime]$validfrom = $dialoginput.validfrom} else {$validfrom = $null}
if ($dialoginput.validto) 	{[DateTime]$validto = $dialoginput.validto} 	else    {$validto = $null}
$sendto = $dialoginput.sendto

Write-AGPLLog "I" "AdaxesCreatePWPDFAndSend" "PW list requested: $validfrom, $validto, $department, $filter" "Initiator: $sendto"

if (-not $sendto) {
	Write-AGPLLog "I" "AdaxesCreatePWPDFAndSend" "No target email defined. Exiting." ""
} else
{

	# check dates
	if (($dialoginput.validfrom) -and ($dialoginput.validto)){
	    if ($validfrom -gt $validto){
	        $tempdate = $validto
	        $validto = $validfrom
	        $validfrom = $tempdate
	    }
	}

	# generate the list of accounts based on parameters
	$listofaccounts = Get-AGPLPasswordLettersDataFromDB -Startdate "$validfrom" -EndDate "$validto" -Filter $filter -Caller $sendto

	# check cases
	if ($listofaccounts.AccountCount -gt 2000)
	{
	    # report return value of call
		Write-AGPLDebug "Your search returned to many results ($($listofaccounts.AccountCount)). Please perform a more specific search. Message: $($listofaccounts.Error)."
	} elseif ($listofaccounts.AccountCount -eq 0) {
	    # report return value of call
		Write-AGPLDebug "Your search did not return any results. Message: $($listofaccounts.Error)"
	} else {
	    # we have some pdfs to generate, do this asynchronously  
	    # first define script block
	    $runasynchron = {
	        param($listofaccounts, [string]$sendto)
	        
	        # load needed modules
	        Import-Module AGPLGeneratePWLetters
	        Import-Module AGPLLogging

	        Generate-AGPLCreatePWLettersAndSend -listOfAccounts $listofaccounts -sendTo $sendto -SMTPServer 'smtp.zhaw.ch'
	    }  # end script block
	    
	    # start creating and sending the pdf async
	    start-job $runasynchron -Arg $listofaccounts,$sendto
	    
	    # report return value of call
	    Write-AGPLDebug "($($listofaccounts.AccountCount)) were found. Producing the requested PDF may take a while. You will receive the resulting PDF by email. This may take a moment. Message: " + $listofaccounts.Error + "."
	}
}