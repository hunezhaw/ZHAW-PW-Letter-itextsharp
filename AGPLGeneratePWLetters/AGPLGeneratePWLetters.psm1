#######################################################################
#
# AGPL License Header
#
#######################################################################
#
# This program is free software; you can redistribute it and/or modify it under the 
# terms of the GNU Affero General Public License version 3 as published by the Free 
# Software Foundation with the addition of the following permission added to Section 
# 15 as permitted in Section 7(a): FOR ANY PART OF THE COVERED WORK IN WHICH THE 
# COPYRIGHT IS OWNED BY 1T3XT, 1T3XT DISCLAIMS THE WARRANTY OF NON INFRINGEMENT OF 
# THIRD PARTY RIGHTS.
#
# This program is distributed in the hope that it will be useful, but WITHOUT ANY 
# WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A 
# PARTICULAR PURPOSE. See the GNU Affero General Public License for more details. 
# You should have received a copy of the GNU Affero General Public License along 
# with this program; if not, see http://www.gnu.org/licenses or write to the Free 
# Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA, 02110-1301 
# USA, or download the license from the following URL: http://itextpdf.com/terms-of-use/
#
# The interactive user interfaces in modified source and object code versions of this 
# program must display Appropriate Legal Notices, as required under Section 5 of the 
# GNU Affero General Public License.
#
# In accordance with Section 7(b) of the GNU Affero General Public License, you must 
# retain the producer line in every PDF that is created or manipulated using iText.
#
# You can be released from the requirements of the license by purchasing a commercial 
# license. Buying such a license is mandatory as soon as you develop commercial 
# activities involving the iText software without disclosing the source code of your 
# own applications. These activities include: offering paid services to customers as 
# an ASP, serving PDFs on the fly in a web application, shipping iText with a closed 
# source product.
#
#######################################################################
#
# Error handling
#
#######################################################################
$ErrorActionPreference = 'Stop'
trap{
	$ErrorMsg = $($error[0] | out-string)
	Write-Host "Error: $ErrorMsg:$($_.InvocationInfo.ScriptName)($($_.InvocationInfo.ScriptLineNumber)): $($_.InvocationInfo.Line)"
	Write-AGPLLog "E" "20000" "psPWLetter" "General error" $ErrorMsg+"$($_.InvocationInfo.ScriptName)($($_.InvocationInfo.ScriptLineNumber)): $($_.InvocationInfo.Line)"
	$errorStat += 1
	continue
}

# load needed modules
try{
	import-module ActiveDirectory
	import-module AGPLLogging
} catch {
	Write-AGPLLog "E" "Generate-AGPLPasswordLetters" "Cannot load dependancies" $_.Exception.Message
	$retObj.Error = "Cannot load dependancies. " + $_.Exception.Message
	$retObj.IsSuccessful = $false
	return $retObj
}

# Definition of access rights groups, % means all
$script:PWLetterRights = @()
$PWLetterRights += New-Object PSObject -Prop @{'GroupName'='Domain Users'	;'Department'='%'	;'Right'='%'	}
$PWLetterRights += New-Object PSObject -Prop @{'GroupName'='Group01'		;'Department'='X'	;'Right'='LE'	}
$PWLetterRights += New-Object PSObject -Prop @{'GroupName'='Group02'		;'Department'='X'	;'Right'='WB'	}
$PWLetterRights += New-Object PSObject -Prop @{'GroupName'='Group03'		;'Department'='X'	;'Right'='STAFF'	}

# Get base path
$script:ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path

# some global variables
$script:encryptionKey = 'YOUR-PW-DECRYPTION-KEY'
$script:sqlConnectionString = $("Data Source=localhost; Database=AUMPWLetterStoreV2AGPL2; User ID=AGPLTestUser; Password=Test.12345")

#######################################################################
#
# Function: getPDFAndFill
# Scope: private
#
# This function gets the defined PDF template,identifies all AcroFields
# and fills them based on the passed AccountInfo.
#
# Author: hune
# Date: 22.03.2013
#
#######################################################################
function script:getPDFAndFill([string]$pdfTemplate, $AccountInfo){
	try{
		$reader = New-Object iTextSharp.text.pdf.pdfreader($pdfTemplate)
		$ms = New-Object System.IO.MemoryStream
		$stamper = New-Object iTextSharp.text.pdf.pdfstamper($reader, $ms)
	} catch {
		Write-AGPLLog "E" "Generate-AGPLPasswordLetters" "Cannot load iTextSharp assembly." $_.Exception.Message
		return $null
	}

	# Make sure, that the PDF document is minimized after filling in content
	[void]$stamper.SetFullCompression()
	[void]$reader.RemoveUnusedObjects

	# set content of form fileds
	$form = $stamper.AcroFields
	$keys = $form.Fields.Keys
	
	Write-AGPLDebug "getPDFAndFill: Starting to fill the PDF."
	
	# iterate through all available fileds and define value
	$value = ''
	foreach($key in $keys){
		switch($Key){
			'ID' 				{$value = $AccountInfo.UniqueID}
			'Name' 				{$value = $AccountInfo.Nachname + ' ' + $AccountInfo.Vorname}
			'CurrentDate'		{$value = $(Get-Date -Format dd.MM.yyyy)}
			'LastName' 			{$value = $AccountInfo.Nachname}
			'FirstName' 		{$value = $AccountInfo.Vorname}
			'SAMAccountName'	{$value = $AccountInfo.KurzZeichen}
			'Class'				{$value = $AccountInfo.AnlassEvento}
			'Password' 			{
									if ($AccountInfo.PWIsValid){
										$value = $AccountInfo.PasswordDecrypted
									} else {
										Write-AGPLDebug "getPDFAndFill: $($AccountInfo.KurzZeichen), PW already changed: $($AccountInfo.PasswordDecrypted), LastSet: $([DateTime]$AccountInfo.PwdLastSet), Timestamp: $([DateTime]$AccountInfo.TimeStamp)"
										$value = 'The password was aready changed by the user!'
									}
								}
			default {
				$value = 'UNKNOWN'		
				Write-AGPLLog "W" "psPWLetter" "Found unknown AcroField in template PDF" $key
			}
		}
		$result = $form.SetField($key,$value)
		Write-AGPLDebug "Looking for key $($key) in PDF, assigning value: $($value)"
	}
    # finish and clean up
	$stamper.Writer.CloseStream = $false;
	$stamper.FormFlattening = $true;
	[void]$stamper.Close()
	[void]$reader.Close()
	$stream = $ms.ToArray()
	
	$returnobj = New-Object iTextSharp.text.pdf.pdfreader(,$stream)
	
	return $returnobj
}

#######################################################################
#
# Function: isUserMemberOfGroup
# Scope: private
#
# This function tests if the user is member of the passed group.
# Grups are checked recursively.
#
# Author: hune
# Date: 08.05.2013
#
#######################################################################
function script:isUserMemberOfGroup([string]$UserName, [string]$GroupName){
	if (!$UserName -or !$GroupName){
		return $false
	}

	try{	
		$objADUser = Get-ADUser $UserName -Properties *
		$objADGroup = Get-ADGroup $GroupName

		if ($objADGroup.DistinguishedName -eq $objADUser.PrimaryGroup) {
			# The requested group is the primary group of the user
			$objResult = Get-ADUser $UserName
		} else {
			# Not the primary group, we have to search some more
			$objResult = Get-ADUser -Filter { memberOf -RecursiveMatch $objADGroup.DistinguishedName } -SearchBase $objADUser.DistinguishedName -SearchScope Base 
		}
	} catch {
		$objResult = $null
		Write-AGPLLog "E" "isUserMemberOfGroup" "Checking a user or group that does not exist" $_.Exception.Message
	}

	if ($objResult) {
		return $true
	} else {
		return $false
	}
}

#######################################################################
#
# Function: getuserQuery
# Scope: private
#
# This function creates the DB query based on the users rights, means
# group memberships.
#
# Author: hune
# Date: 07.05.2013
#
#######################################################################
function script:getUserQuery([string]$UserName){

	if (!$UserName){
		return ''
	}

	# define return object
	$retObj = New-Object Object
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'SQLUserQuery' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'UserRights' -Value ''
	
	# define variables
	$found = $false
	$RightAll = $false
	$RightStaff = $false
	$RightLE = $false
	$RightWB = $false
	
	# define the base SQL table variable
	$SQLUserQuery = 'DECLARE @DepDataRights AS typ_DepDataRights;'
	
	# now based on the users group membership, add the needed entries to this table
	foreach ($right in $PWLetterRights){
		if (isUserMemberOfGroup $UserName $right.GroupName){
			$found = $true
			Write-AGPLDebug "User $UserName is member of group $($right.GroupName)"
			$SQLUserQuery += "INSERT INTO @DepDataRights (Department, DataRight) VALUES ('$($right.Department)', '$($right.Right)');"
			
			switch ($right.Right){
				'LE'	{
							$RightLE = $true
						}
				'WB'	{
							$RightWB = $true
						}
				'STAFF'	{
							$RightStaff = $true
						}
				'%'		{
							$RightAll = $true
						}
			}
		}
	}
	
	# translate the group rights into a string we can use for a file name
	if ($RightLE) {
		$retObj.UserRights += 'LE'
	}
	if ($RightWB) {
		$retObj.UserRights += 'WB'
	}
	if ($RightStaff) {
		$retObj.UserRights = 'STAFF'
	}
	if ($RightAll -or (($RightLE -or $RightWB) -and $RightStaff)) {
		$retObj.UserRights = 'ALL'
	}
	$retObj.SQLUserQuery = $SQLUserQuery

	# return the data. if the user does not have rights, return $null.
	if ($found) {
		return $retObj
	} else {
		return $null
	}
}

<#
	.SYNOPSIS
	Get full name of file in base directory
	.DESCRIPTION
	Allows accessing some files in the base directory. Returns the full path and file name. 
	If file not there, an empty string is returned. Used for accessing template files.
	.PARAMETER FileName
	Name of file to be accessed
	.EXAMPLE
	Get-AGPLPasswordLetterExpandFileName 'Passwortbrief_DEMO.pdf'
	.NOTES
#>
function Get-AGPLPasswordLetterExpandFileName([string]$FileName=$null){
	$fullFileName = Join-Path $ScriptDirectory $FileName
	if (Test-Path $fullFileName){
		return $fullFileName
	} else {
		return ''
	}
}

<#
	.SYNOPSIS
	Prepare list of PW letters to be converted to PDF
	.DESCRIPTION
	Reads information from the password store, based on the passed parameters and creates an array of elements to be further processed
	.PARAMETER StartDate
	Start date of periode. Optional. If not set, User must be set.
	.PARAMETER EndDate
	End date of periode. Optional, current date if not set.
	.PARAMETER Filter
	A filter for limiting the returned PWLetters, searches in Vorname, Nachname, KurzZeichen, Klasse, Departement
	.PARAMETER User
	User short name (Kürzel). If present all other parameters are omitted. Only one pw letter is created.
	.PARAMETER Caller
	User short name (Kürzel). Used to define the scope of the returned data.
	.EXAMPLE
	Get-AGPLPasswordLettersDataFromDB -StartDate '27.03.2013' -EndDate '28.03.2013' -Caller 'hune'
	.NOTES
#>
function Get-AGPLPasswordLettersDataFromDB([string]$StartDate=$null,[string]$EndDate=$null,[string]$Filter=$null,[string]$User=$null,[parameter(mandatory=$true)][string]$Caller){	
	# generate return object
	$retObj = New-Object Object
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'AccountCount' -Value 0
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'ElementList' -Value $null
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'UserRights' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'SearchParameter' -Value ''	
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'Caller' -Value $Caller	
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'Error' -Value 'Operation successful!'
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'IsSuccessful' -Value $true
	
	# handle different date format, we need a nice text for the mail to be sent
	$localStartDate = 'No Startdate'
	$localEndDate = 'No Enddate'

	try{
		If ($StartDate) {
			If ($StartDate.Contains('/')) {
				$localStartDate = [datetime]::ParseExact($StartDate.Substring(0,10), "MM'/'dd'/'yyyy",$null).toString("dd.MM.yyyy")
				
			} else {
				$localStartDate = [datetime]::ParseExact($StartDate.Substring(0,10), "dd.MM.yyyy",$null).toString("dd.MM.yyyy")
			}
		}
		If ($EndDate) {
			If ($EndDate.Contains('/')) {
				$localEndDate = [datetime]::ParseExact($EndDate.Substring(0,10), "MM'/'dd'/'yyyy",$null).toString("dd.MM.yyyy")
			} else {
				$localEndDate = [datetime]::ParseExact($EndDate.Substring(0,10), "dd.MM.yyyy",$null).toString("dd.MM.yyyy")
			}
		}
	} catch {
		$StartDate = $null
		$EndDate = $null
		$localStartDate = 'No Startdate'
		$localEndDate = 'No Enddate'
	}
	$retObj.SearchParameter = "$($localStartDate), $($localEndDate), $(if ($filter) {$filter} else {'No Filter'})"
	
	# initialize
	$errorStat		 		= 0
	$sqlQuery				= ''

	# validate parameters
	$SQLStartDate 		= if($StartDate){"'$StartDate'"}else{'NULL'}
	$SQLEndDate 		= if($EndDate){"'$EndDate'"}else{'NULL'}
	$SQLKey 			= "'$encryptionKey'"
	$SQLUser 			= if($User){"'$User'"}else{'NULL'}
		
	# search for limiting patterns. if this patterns are found, limit the search result
	$departments 	= '{A}','{G}','{L}','{N}','{P}','{R}','{S}','{T}','{V}','{W}','{X}'
	$types 			= '{WB}','{LE}','{STAFF}'

	# search for department pattern
	$departmentList = $null
	$typeList = $null
	if ($filter) {
		foreach ($department in $departments)
		{ 
			if ($filter -match $department) {
				$filter = $filter -replace $department, ''
				if ($departmentList) {
					$departmentList += ',' + $department.Substring(1,1)
				} else {
					$departmentList = $department.Substring(1,1)
				}
			}	
		}
	}

	# search for type pattern
	if ($filter) {
		foreach ($type in $types)
		{ 
			if ($filter -match $type) {
				$filter = $filter -replace $type, ''
				if ($typeList) {
					$typeList += ',' + $type.Substring(1,$type.length - 2)
				} else {
					$typeList = $type.Substring(1,$type.length - 2)
				}
			}	
		}
	}
	if (!$departmentList) {$departmentList = 'NULL'} else {$departmentList = "'$departmentList'"}
	if (!$typeList) {$typeList = 'NULL'} else {$typeList = "'$typeList'"}
	$SQLFilter	= if($Filter){"'$Filter'"}else{'NULL'}
	$SQLFilter = $SQLFilter -replace "\*", "%" # Map user search token to SQL search token
	$SQLFilter = $SQLFilter -replace "\?", "_" # Map user search token to SQL search token

	Write-AGPLDebug "Get-AGPLPasswordLettersDataFromDB:Requesting user: $Caller"
	
	# prepare query
	if ($User) {
		# we are getting PW letter of one user
		$sqlQuery = "exec usp_PWStoreGetUserData @User=$SQLUser,@Key=$SQLKey"
		$sqlUserRights = $user
	} else {
		# this is a request for multiple PW letters
		# was a user name passed to the call? we need this to define what the user can see
		if ($Caller) {
			# clean up caller
			if ($Caller.split('\').Count -gt 1) {
				# we have domain\name
				$Caller = $Caller.split('\')[-1]
			} elseif ($Caller.split('@').Count -gt 1) {
				# we have name@domain
				$Caller = $Caller.split('@')[0]
			}		
		
			$userQueryResult = getUserQuery $Caller
			if ($userQueryResult) {
				# Put together the final call to the stored procedure
				$sqlQuery = $userQueryResult.SQLUserQuery
				$sqlUserRights = $userQueryResult.UserRights
				$sqlQuery += "exec usp_PWStoreGetData @startDate=$SQLStartDate,@endDate=$SQLEndDate,@Key=$SQLKey,@Filter=$SQLFilter,@Rights=@DepDataRights,@Department=$departmentList,@Type=$typeList"	
			} else {
				Write-AGPLLog "E" "GGet-AGPLPasswordLettersDataFromDB" "No rights to access data." 
				$retObj.Error = "No rights to access data. "
				$retObj.IsSuccessful = $false
				return $retObj
			}
		} else {
			Write-AGPLLog "E" "Get-AGPLPasswordLettersDataFromDB" "No caller defined. Cannot create user query for DB." 
			$retObj.Error = "No caller defined. Cannot create user query for DB. "
			$retObj.IsSuccessful = $false
			return $retObj
		}
	}
	Write-AGPLDebug "Get-AGPLPasswordLettersDataFromDB:Used SQL query: $sqlQuery"

	# make DB stuff
	try{
		$Table = new-object System.Data.DataTable
		$sqlConn = new-object System.Data.SqlClient.SqlConnection($sqlConnectionString)
		$adapter = new-object System.Data.SqlClient.SqlDataAdapter($sqlQuery,$sqlConn)
		$tableCount = $adapter.Fill($Table)
	} catch {
		Write-AGPLLog "E" "Get-AGPLPasswordLettersDataFromDB" "Error accessing DB" $_.Exception.Message
		Write-AGPLLog "E" "Get-AGPLPasswordLettersDataFromDB" "DB Query" $sqlQuery
		$retObj.Error = 'Error connecting to DB. ' + $_.Exception.Message
		$retObj.IsSuccessful = $false
		return $retObj
	}

	Write-AGPLDebug "Get-AGPLPasswordLettersDataFromDB: Number of elements from DB: $($tableCount)"
	
	# prepare return object
	if ($tableCount -gt 0){
		$retObj.ElementList = $Table.Select()
		$retObj.UserRights = $sqlUserRights
		$retObj.AccountCount = $tableCount
		$retObj.IsSuccessful = $true
	} else {
		$retObj.ElementList = $null
		$retobj.UserRights = ''
		$retObj.AccountCount = 0
		$retObj.IsSuccessful = $true
		$retObj.Error = 'The search did not return any content!'
	}
	
	# return values
	return $retObj
}

<#
	.SYNOPSIS
	Creation of password letters from passed data
	.DESCRIPTION
	Creates the requested password letters (one file)
	.PARAMETER inputObject
	Object returned by a call to Get-AGPLPasswordLettersDataFromDB
	.EXAMPLE
	Generate-AGPLPasswordLettersFromList -inputObject $resultFromDB
	.NOTES
#>
function Generate-AGPLPasswordLettersFromList($inputObject){	
	# generate return object
	$retObj = New-Object Object
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'PageCount' -Value 0
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'AccountCount' -Value 0
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'OutputFileName' -Value ''
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'Error' -Value 'Operation successful!'
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'IsSuccessful' -Value $true
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'BasePath' -Value $ScriptDirectory

	# init variables
	if ($inputObject -and $inputObject.PSObject.Properties.Match('ElementList') -and $inputObject.ElementList){
		$extendedElements = $inputObject.ElementList
		$userRights = $inputObject.UserRights
	} else {
		Write-AGPLLog "E" "Generate-AGPLPasswordLettersFromList" "Error in paramter object"
		$retObj.Error = 'Error in paramter to Generate-AGPLPasswordLettersFromList'
		$retObj.IsSuccessful = $false
		return $retObj
	}

	# load the pdf library
	try{
		$iTextSharpAssembly = Join-Path $ScriptDirectory "itextsharp\itextsharp.dll"
		[void][System.Reflection.Assembly]::LoadFrom($iTextSharpAssembly) 
	} catch {
		Write-AGPLLog "E" "Generate-AGPLPasswordLettersFromList" "Cannot load itextsharp.dll" $_.Exception.Message
		$retObj.Error = 'Cannot load itextsharp.dll. ' + $_.Exception.Message
		$retObj.IsSuccessful = $false
		return $retObj
	}

	Write-AGPLDebug "Generate-AGPLPasswordLettersFromList: Number of elements: $($extendedElements.Count)"
	
	if ($extendedElements.Count -gt 0){
		# prepare output document
		$outputPath = Join-Path $ScriptDirectory "PDFOutput"
		if (!(Test-Path -path $outputPath)) {$outputPath = New-Item $outputPath -Type Directory}

		# Put togehter the correct file name (based on the user rights returned by getUserQuery)
		$outPDFFile = Join-Path $outputPath "%DATE%-PW-$($userRights).pdf"
		$outPDFFile = $outPDFFile -replace "%DATE%", $([datetime]::Now).ToString("hhmmss-ddMMyyyy")

		Write-AGPLDebug "Generate-AGPLPasswordLettersFromList: Starting PDF production"
		
		# start filling content and produce output pdf
		try{
			$doc = New-Object itextsharp.text.document
			$stream = [IO.File]::OpenWrite($outPDFFile)
			$writer = New-Object itextsharp.text.pdf.PdfSmartCopy($doc, $stream)
			[void]$writer.SetFullCompression()
			[void]$doc.Open()
		} catch {
			Write-AGPLLog "E" "Generate-AGPLPasswordLettersFromList" "Error opening output file: " + $outPDFFile $_.Exception.Message
			$retObj.Error = 'Cannot open templates. ' + $_.Exception.Message
			$retObj.IsSuccessful = $false
			return $retObj
		}
		
		Write-AGPLDebug "Generate-AGPLPasswordLettersFromList: Looping through elements"
		
		# loop through the elements and create the pdf files
		$pageCount = 0
		foreach ($element in $extendedElements){
			try{
				[void]$doc.NewPage()
				# define the needed template file name
				if (($element.AccountType -eq 'LE') -or ($element.AccountType -eq 'WB')){
					$inPDFFile = Join-Path $ScriptDirectory "ExamplePDF-01.pdf"		
				} else {
					$inPDFFile = Join-Path $ScriptDirectory "ExamplePDF-02.pdf"		
				}
				$pdfReader = getPDFAndFill $inPDFFile $element
				[void]$pdfReader.RemoveUnusedObjects
				
				[iTextSharp.text.pdf.PdfImportedPage] $page = $writer.GetImportedPage($pdfReader, 1)
				[Void]$writer.AddPage($page)
				[iTextSharp.text.pdf.PdfImportedPage] $page = $writer.GetImportedPage($pdfReader, 2)
				[Void]$writer.AddPage($page)
				[Void]$writer.FreeReader($pdfReader)
				[Void]$pdfReader.Close()
				[Void]$stream.Flush()
				$pageCount += 2

				# in order not to fill the RAM on the server we flush to disk every 100 pages
				if ($pageCount % 100 -eq 0){
					[Void]$writer.Flush()
				}
			} catch {
				Write-AGPLLog "E" "Generate-AGPLPasswordLettersFromList" "Error processing account entries" $_.Exception.Message
			}		
		}

		# flush content of writer to disk
		[Void]$writer.Flush()
		
		Write-AGPLDebug "Generate-AGPLPasswordLettersFromList: Finished looping through elements"
		
		# provide some return values
		$retObj.OutputFileName = $outPDFFile
		$retObj.AccountCount = $extendedElements.Count
		$retObj.PageCount = $pageCount
		$retObj.IsSuccessful = $true

		try{
			# Clean up and finalize
			[void]$doc.Close()
			[void]$writer.Close()
		} catch {
			Write-AGPLLog "E" "Generate-AGPLPasswordLettersFromList" "Error closing output file" $_.Exception.Message
		}
	} else {
		$retObj.OutputFileName = $null
		$retObj.AccountCount = 0
		$retObj.PageCount = 0
		$retObj.IsSuccessful = $true
		$retObj.Error = 'The search did not return any content!'
	}
	
	# return values
	return $retObj
}

<#
	.SYNOPSIS
	Calls Generate-AGPLPasswordLettersFromList and sends PWLetters by mail
	.DESCRIPTION
	This function can be called async.
	.PARAMETER listofaccounts
	Result object from Get-AGPLPasswordLettersDataFromDB, containing the list of PWLetters to produce
	.PARAMETER sendto
	Target EMail account
	.EXAMPLE
	Generate-AGPLCreatePWLettersAndSend -listOfAccounts $listofaccounts -sendTo 'hune@AGPL.ch'
	.NOTES
#>
function Generate-AGPLCreatePWLettersAndSend($listOfAccounts, [string]$sendTo, [string]$SMTPServer)
{	       
    Write-AGPLDebug "Generate-AGPLCreatePWLettersAndSend: Start producing PDF"
	$startDate = Get-Date
	
	# generate the pdf, may be long running
    $generatepasswordletters = Generate-AGPLPasswordLettersFromList -inputObject $listOfAccounts
    
    if (($generatepasswordletters.IsSuccessful) -and ($generatepasswordletters.Pagecount -gt 0)){
        # success, send out the requested PDF file
        $bodyContentPath = Join-Path $ScriptDirectory "MailBodyPWLetterRequest.html"
        $Body = Get-Content $bodyContentPath | out-string
        $Encoding  = New-Object System.Text.utf8encoding
    
        $dateFormat = New-Object System.Globalization.CultureInfo("de-CH")
        $Body = $Body -replace 'ACCOUNTNUMBER', $listofaccounts.AccountCount
        $Body = $Body -replace 'SEARCHPARAMETER', $listofaccounts.SearchParameter
    
        send-mailmessage –BodyAsHtml -Encoding $Encoding  -from "dont.reply@AGPL.ch" -to $sendTo -subject "Your generated list of letters" -body $Body -Attachments "$($generatepasswordletters.OutputFileName)" -SmtpServer $SMTPServer
    } else {
        # fail, send out information
        $bodyContentPath = Join-Path $ScriptDirectory "MailBodyPWLetterRequestFail.html"
        $Body = Get-Content $bodyContentPath | out-string
        $Encoding  = New-Object System.Text.utf8encoding
    
        $dateFormat = New-Object System.Globalization.CultureInfo("de-CH")
        $Body = $Body -replace 'SEARCHPARAMETER', $listofaccounts.SearchParameter
        $Body = $Body -replace 'ERRORMESSAGE', $generatepasswordletters.Error
    
        send-mailmessage –BodyAsHtml -Encoding $Encoding  -from "dont.reply@AGPL.ch" -to $sendTo -subject "Your letters could not be created" -body $Body -SmtpServer "smtp.AGPL.ch"
    } 
	
	Write-AGPLDebug "Generate-AGPLCreatePWLettersAndSend: End producing PDF: $((new-TimeSpan $startDate $(get-date)).TotalSeconds) seconds"
}

<#
	.SYNOPSIS
	Writes a password record to the PW store
	.DESCRIPTION
	Base functionality to add a PW record to the password store.
	.PARAMETER parameters
	A hash table with the records information. All the following fileds must be present in the hash ('UniqueID','KurzZeichen','PersKat','Nachname','Vorname','Geschlecht','EMail1','EMail2','Departement','Vorgesetzter','DepDescription','Password')
	.EXAMPLE
	Write-AGPLPasswordToStore $parameters 
	.NOTES
#>
function Write-AGPLPasswordToStore([hashtable]$parameters)
{
	# generate return object
	$isError = $false
	$retObj = New-Object Object
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'Error' -Value 'Operation successful!'
	Add-Member -InputObject $retObj -MemberType NoteProperty -Name 'IsSuccessful' -Value $true
	
	# prepare query. attention timestamp plus 1 minutes because of synchronisation
	$properties = @('UniqueID','KurzZeichen','Departement','DepDescription','Password')
	$sqlQuery = "INSERT INTO tPWHistory (TimeStamp,UniqueID,KurzZeichen,Departement,DepDescription,PasswordEncrypted)  
				 VALUES (DATEADD(minute,1,getdate()), 'ValUniqueID', 'ValKurzZeichen', 'ValDepartement', 'ValDepDescription', EncryptByPassPhrase('$encryptionKey','ValPassword'))"
	
	foreach ($property in $properties){
		if ($parameters.ContainsKey($property)){
			$sqlQuery = $sqlQuery.Replace('Val'+$property, $parameters.Get_Item($property))
		} else {
			$isError = $true
			Write-AGPLLog "E" "Write-AGPLPasswordToStore" "Missing key in passed parameters" $property
		}
	}
	
	Write-AGPLDebug $sqlQuery
	
	if (!$isError) {
		# write password entry to store
		try{
			$Table = new-object System.Data.DataTable
			$sqlConn = new-object System.Data.SqlClient.SqlConnection($sqlConnectionString)
			$sqlCommand = new-object System.Data.SqlClient.sqlcommand($sqlQuery,$sqlConn)
			[Void]$sqlConn.Open()
			[void]$sqlCommand.ExecuteNonQuery()
		} catch {
			Write-AGPLLog "E" "Write-AGPLPasswordToStore" "Error accessing DB" $_.Exception.Message
			Write-AGPLLog "E" "Write-AGPLPasswordToStore" "DB Query" $sqlQuery
			$retObj.Error = 'Error executing query. ' + $_.Exception.Message
			$retObj.IsSuccessful = $false
		}
	} else {
		Write-AGPLLog "E" "Write-AGPLPasswordToStore" "Parameter error. No password inserted." ""
		$retObj.Error = 'Parameter error. No password inserted.'
		$retObj.IsSuccessful = $false
	}
	
	return $retObj
}

Export-ModuleMember -function Get-AGPLPasswordLettersDataFromDB
Export-ModuleMember -function Generate-AGPLCreatePWLettersAndSend
Export-ModuleMember -function Generate-AGPLPasswordLettersFromList
Export-ModuleMember -function Get-AGPLPasswordLetterExpandFileName
Export-ModuleMember -function Write-AGPLPasswordToStore




