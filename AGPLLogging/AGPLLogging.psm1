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
	Write-Host "Error: $ErrorMsg"
	Write-AGPLLog "E" "20000" "AGPLLogging" "General error" $ErrorMsg
	$errorStat += 1
	continue
}

# Definition of logging constants
set-variable -name constDebugOutput -value $TRUE -option constant
set-variable -name constDebugIncLog -value $FALSE -option constant
set-variable -name constDebugFile -value "%DATE%-Debug.txt" -option constant
set-variable -name constDebugLogFile -value "%DATE%-Log.txt" -option constant

# Get base path
$script:ScriptDirectory = Split-Path $MyInvocation.MyCommand.Path

<#
	.SYNOPSIS
	Writes a log entry to the log file.
	.DESCRIPTION
	Writes a log entry to the log file.
	.PARAMETER type
	Type of the log entry. Should be E, W, I.
	.PARAMETER source
	Name of the script writing the log entry.
	.PARAMETER text1
	First text part of the log.
	.PARAMETER text2
	Second text part of the log.
	.PARAMETER text3
	Third text part of the log.
	.EXAMPLE
	Write-AGPLLog "E" "Generate-AGPLPasswordLetters" "Error closing output file" $_.Exception.Message 
	.NOTES
#>
# Function Write-AGPLLog
Function Write-AGPLLog([string]$type, [string]$source, [string]$text1, [string]$text2, [string]$text3){
	
	$timeStamp = Get-Date
	$message = [string]::Format('{0};{1};{2};{3};{4};{5};{6}',$($timeStamp.ToString('dd.MM.yyyy')),$($timeStamp.ToLongTimeString()),$Type,$Source,$Text1,$Text2,$Text3)

	# Write message
	if ($constDebugIncLog){
		Write-AGPLDebug $message
	}

	$outputPath = Join-Path $ScriptDirectory "LogFiles"
	if (!(Test-Path -path $outputPath)) {New-Item $outputPath -Type Directory}	
		
	$fileName = $constDebugLogFile -replace "%DATE%", $([datetime]::Now.Date).ToString("ddMMyyyy")
	$fileName = Join-Path $outputPath $fileName
	$stream = new-object System.IO.StreamWriter($fileName, $true)
	$stream.WriteLine($message)
	$stream.close()
}

<#
	.SYNOPSIS
	Writes a log entry to the debug file.
	.DESCRIPTION
	Writes a log entry to the debug file.
	.PARAMETER message
	The message to be written to the debug file.
	.EXAMPLE
	Write-AGPLDebug "Test debug message"
	.NOTES
#>
# Function Write-AGPLDebug
Function Write-AGPLDebug([string]$message){
	# Write message
	$timeStamp = Get-Date	
	$fileName = $constDebugFile -replace "%DATE%", $([datetime]::Now.Date).ToString("ddMMyyyy")
	
	$outputPath = Join-Path $ScriptDirectory "LogFiles"
	if (!(Test-Path -path $outputPath)) {New-Item $outputPath -Type Directory}	
	
	$fileName = Join-Path $outputPath $fileName
	$msg = [string]::Format('{0};{1};{2}',$($timeStamp.ToString('dd.MM.yyyy')),$($timeStamp.ToLongTimeString()),$message)
	$stream = new-object System.IO.StreamWriter($fileName, $true)	
	$stream.WriteLine($msg)
	$stream.close()
}

# Export members
Export-ModuleMember -function Write-AGPLLog
Export-ModuleMember -function Write-AGPLDebug