# ----------------------------------------------------------------------------
#	Licensed under the Apache License, Version 2.0 (the "License");
#	you may not use this file except in compliance with the License.
#	You may obtain a copy of the License at
#
#		http://www.apache.org/licenses/LICENSE-2.0
#
#	Unless required by applicable law or agreed to in writing, software
#	distributed under the License is distributed on an "AS IS" BASIS,
#	WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#	See the License for the specific language governing permissions and
#	limitations under the License.
# ----------------------------------------------------------------------------
#
#	Copyright (c) 2020 raphael.klein@gmail.com  All rights reserved.
#
# title			Outlook Signature from Template for AD Users
# description	A PowerShell script that creates Outlook Signatures from 
#				Templates using AD Data.
#				available and calls the subscript chosen. 
# author		raphael.klein@gmail.com
# usage			1) Place .docx Template(s) into \\My_Domain\NETLOGON\signature_template\
#				2) Execute Script
#				   PS > powershell.exe OutlookSignatureFromTemplate.ps1
# notes			Run on every user login using task scheduler:
# 				PS > $Trigger= New-ScheduledTaskTrigger -AtStartup
# 				PS > $User= Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -expand UserName
# 				PS > $Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument `
# 				PS > 	“-File C:\PS\OutlookSignatureFromTemplate.ps1”
# 				PS > Register-ScheduledTask -TaskName "CreateOutlookSignatureFromTemplate" `
# 				PS > 	-Trigger $Trigger -User $User -Action $Action
# todo			* Notifications (System.Windows.Forms.NotifyIcon)
# ----------------------------------------------------------------------------

# Only if AD available
if (!(Test-Connection -ComputerName (Get-WmiObject Win32_ComputerSystem).Domain -Quiet)) {
	Write-Host "This script requires a connection to the Domain Controller, try again when you are connected."; Exit
}

# $TemplateFolderPath = "$($env:TEMP)\signature_template\"
$TemplateFolderPath = "\\$((Get-WmiObject Win32_ComputerSystem).Domain)\NETLOGON\signature_template\"
$DestinationFolderPath = (Get-Item env:appdata).value+"\Microsoft\Signatures\"

$Searcher = New-Object system.directoryservices.directorysearcher “samAccountName=$env:username”
$User = $Searcher.FindOne().GetDirectoryEntry()

$FIRSTNAME = If ( "givenName" -in $User.PSobject.Properties.Name ) { $User.givenName } Else { "" }
$LASTNAME = If ( "sn" -in $User.PSobject.Properties.Name ) { $User.sn } Else { "" }
$EMAIL = If ( "mail" -in $User.PSobject.Properties.Name ) { $User.mail } Else { "" }
$ADDRESS = If ( "streetAddress" -in $User.PSobject.Properties.Name ) { $User.streetAddress } Else { "" }
If ( $ADDRESS -and "postalCode" -in $User.PSobject.Properties.Name ) { $ADDRESS = "$($ADDRESS), $($User.postalCode)" }
If ( $ADDRESS -and "l" -in $User.PSobject.Properties.Name ) { $ADDRESS = "$($ADDRESS) $($User.l)" }
If ( $ADDRESS -and "co" -in $User.PSobject.Properties.Name ) { $ADDRESS = "$($ADDRESS), $($User.co)" }
$COMPANY = If ( "company" -in $User.PSobject.Properties.Name ) { $User.company } Else { "" }
$DEPARTMENT = If ( "department" -in $User.PSobject.Properties.Name ) { $User.department } Else { "" }
$TITLE = If ( "title" -in $User.PSobject.Properties.Name ) { $User.title } Else { "" }
$MOBILENUMBER = If ( "mobile" -in $User.PSobject.Properties.Name ) { $User.mobile } Else { "" }
$FIXEDLINENUMBER = If ( "telephoneNumber" -in $User.PSobject.Properties.Name ) { $User.telephoneNumber } Else { "" }
$TEAMNAME = If ( "team" -in $User.PSobject.Properties.Name ) { $User.team } Else { "" }

Get-ChildItem -Path "$($TemplateFolderPath)*" -Include *.docx  |
Foreach-Object {
	$TemplateFullPath = $_.FullName
	$DestinationFullPath = "$($DestinationFolderPath)$($_.BaseName).rtf"
	
	# Skip if there is no newer signature available
	if (Test-Path $DestinationFullPath) {
		if ((Get-ItemProperty -Path $TemplateFullPath).LastWriteTime -lt (Get-ItemProperty -Path $DestinationFullPath).LastWriteTime) {
			Write-Host "No update required"; return
		}
	}
	
	add-type -AssemblyName “Microsoft.Office.Interop.Word”
	$wdunits = “Microsoft.Office.Interop.Word.wdunits” -as [type]

	$WordInstance = New-Object -ComObject Word.Application
	$WordInstance.Visible = $false
	$WordDocument = $WordInstance.Documents.Open($TemplateFullPath, $false, $true)
	
	$range = $WordDocument.Content
	$range.movestart($wdunits::wdword,$range.start) | Out-Null
	
	If($FIRSTNAME -ne "") {$ReplaceText = $FIRSTNAME.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("FIRSTNAME", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($LASTNAME -ne "") {$ReplaceText = $LASTNAME.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("LASTNAME", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($EMAIL -ne "") {$ReplaceText = $EMAIL.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("EMAIL", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	if ($range.find.execute($ReplaceText,$true,$true,$false,$false,$false,$true,1)) {
		if($range.style.namelocal -eq “normal”) {
			$WordDocument.HyperLinks.Add($range, "mailto:"+$ReplaceText.ToString(),$null,$null,$ReplaceText.ToString()) | Out-Null
		}
	}
	$range.movestart($wdunits::wdword,$range.start) | Out-Null
	$wordFound = $false
	
	If($ADDRESS -ne "") {$ReplaceText = $ADDRESS.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("ADDRESS", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($COMPANY -ne "") {$ReplaceText = $COMPANY.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("COMPANY", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($DEPARTMENT -ne "") {$ReplaceText = $DEPARTMENT.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("DEPARTMENT", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($TEAMNAME -ne "") {$ReplaceText = $TEAMNAME.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("TEAMNAME", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($TITLE -ne "") {$ReplaceText = $TITLE.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("TITLE", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	
	If($MOBILENUMBER -ne "") {$ReplaceText = $MOBILENUMBER.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("MOBILENUMBER", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	if ($range.find.execute($ReplaceText,$true,$true,$false,$false,$false,$true,1)) {
		if($range.style.namelocal -eq “normal”) {
			$WordDocument.HyperLinks.Add($range, "tel:"+$ReplaceText.ToString(),$null,$null,$ReplaceText.ToString()) | Out-Null
		}
	}
	$range.movestart($wdunits::wdword,$range.start)
	$wordFound = $false
	
	If($FIXEDLINENUMBER -ne "") {$ReplaceText = $FIXEDLINENUMBER.ToString()} Else {$ReplaceText = ""}
	$range.Find.Execute("FIXEDLINENUMBER", $True, $True, $False, $False, $False, $True, 1, $False, $ReplaceText, 2 )
	if ($range.find.execute($ReplaceText,$true,$true,$false,$false,$false,$true,1)) {
		if($range.style.namelocal -eq “normal”) {
			$WordDocument.HyperLinks.Add($range, "tel:"+$ReplaceText.ToString(),$null,$null,$ReplaceText.ToString()) | Out-Null
		}
	}
	$range.movestart($wdunits::wdword,$range.start)
	$wordFound = $false
		
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
	[ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type]

	$WordDocument.WebOptions.OrganizeInFolder = $true
	$WordDocument.WebOptions.UseLongFileNames = $true
	$WordDocument.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6
	$SavePath = "$($DestinationFolderPath)$($_.BaseName).htm"
	$WordDocument.saveas([ref]$SavePath, [ref]$saveFormat)

	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
	$SavePath = "$($DestinationFolderPath)$($_.BaseName).rtf"
	$WordDocument.SaveAs([ref] $SavePath, [ref]$saveFormat)

	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");
	$SavePath = "$($DestinationFolderPath)$($_.BaseName).txt"
	$WordDocument.SaveAs([ref] $SavePath, [ref]$saveFormat)
	$WordInstance.Documents.Close([ref] [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)

	$WordInstance.Quit()
}
