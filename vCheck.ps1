param ([Switch]$config, $Outputpath)
###############################
# vCheck - Daily Error Report # 
###############################
# Thanks to all who have commented on my blog to help improve this project
# all beta testers and previous contributors to this script.
#
# Added code for Database exports of results
$Version = "6.17"
$a = $host.PrivateData
#
# Grab the Warning foreground and background colors the user configured
#
$Warning_Foreground = $a.WarningForegroundColor
$Warning_Background = $a.WarningBackgroundColor

function Write-CustomOut ($Details){
	$LogDate = Get-Date -Format T
	Write-Host "$($LogDate) $Details"
	#write-eventlog -logname Application -source "Windows Error Reporting" -eventID 12345 -entrytype Information -message "vCheck: $Details"
}
Function Get-ID-String ($file_content,$ID_name) {
if ($file_content | Select-String -Pattern "\$+$ID_name\s*=")
  {	
  # Write-CustomOut "${ID_name}: ", $file_content | Select-String -Pattern "\$+${ID_name}\s*="
  	$value = (($file_content | Select-String -pattern "\$+${ID_name}\s*=").toString().split("=")[1]).Trim(' "')
	 return ( $value ) }
}

Function Get-PluginID ($Filename){
# Get the identifying information for a plugin script
# Write-Host "Filename: $Filename"
  $file = Get-Content $Filename
  $Title = Get-ID-String $file "Title"
  if ( !$Title ) { $Title = $Filename }
  $PluginVersion = Get-ID-String $file "PluginVersion"
  $Author = Get-ID-String $file "Author"
  $Ver = "{0:N1}" -f $PluginVersion
			
# Write-Host "Title: $Title, PluginVersion: $PluginVersion, Ver: $Ver, Author: $Author"
Return [array]( $Title, $Ver, $Author )			
}

Function Invoke-Settings ($Filename, $GB) {
	$file = Get-Content $filename
	$OriginalLine = ($file | Select-String -Pattern "# Start of Settings").LineNumber
	$EndLine = ($file | Select-String -Pattern "# End of Settings").LineNumber
	if (($OriginalLine +1) -eq $EndLine) {
		} Else {
		$Array = @()
		$Line = $OriginalLine
		do {
			$Question = $file[$Line]
			$Line ++
			$Split= ($file[$Line]).Split("=")
			$Var = $Split[0]
			$CurSet = $Split[1]
			
			# Check if the current setting is in speach marks
			$String = $false
			if ($CurSet -match '"') {
				$String = $true
				$CurSet = $CurSet.Replace('"', '')
			}
			$NewSet = Read-Host "$Question [$CurSet]"
			If (-not $NewSet) {
				$NewSet = $CurSet
			}
			If ($String) {
				$Array += $Question
				$Array += "$Var=`"$NewSet`""
			} Else {
				$Array += $Question
				$Array += "$Var=$NewSet"
			}
			$Line ++ 
		} Until ( $Line -ge ($EndLine -1) )
		$Array += "# End of Settings"

		$out = @()
		$out = $File[0..($OriginalLine -1)]
		$out += $array
		$out += $File[$Endline..($file.count -1)]
		if ($GB) { $out[$SetupLine] = '$SetupWizard =$False' }
		$out | Out-File $Filename
	}
}

# Add all global variables.
$ScriptPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
$ProviderName = "VMWareVCheck"
$PluginsFolder = $ScriptPath + "\Plugins\"
$Plugins = Get-ChildItem -Path $PluginsFolder -filter "*.ps1" | Sort Name
$GlobalVariables = $ScriptPath + "\GlobalVariables.ps1"




$file = Get-Content $GlobalVariables

$Setup = ($file | Select-String -Pattern '# Set the following to true to enable the setup wizard for first time run').LineNumber
$SetupLine = $Setup ++
$SetupSetting = invoke-Expression (($file[$SetupLine]).Split("="))[1]
if ($config) {
	$SetupSetting = $true
}
If ($SetupSetting) {
	cls
	Write-Host -foreground $Warning_Foreground -background $Warning_Background
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "Welcome to vCheck by Virtu-Al http://virtu-al.net"
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "================================================="
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "This is the first time you have run this script or you have re-enabled the setup wizard."
	Write-Host -foreground $Warning_Foreground -background $Warning_Background
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "To re-run this wizard in the future please use vCheck.ps1 -Config"
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "To define a path to store each vCheck report please use vCheck.ps1 -Outputpath C:\tmp"
	Write-Host -foreground $Warning_Foreground -background $Warning_Background
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "Please complete the following questions or hit Enter to accept the current setting"
	Write-Host -foreground $Warning_Foreground -background $Warning_Background "After completing ths wizard the vCheck report will be displayed on the screen."
	Write-Host -foreground $Warning_Foreground -background $Warning_Background
	
	Invoke-Settings -Filename $GlobalVariables -GB $true
	Foreach ($plugin in $Plugins) { 
		Invoke-Settings -Filename $plugin.Fullname
	}
}

. $GlobalVariables

$vcvars = @("SetupWizard" , "Server" , "SMTPSRV" , "EmailFrom" , "EmailTo" , "EmailSubject", "DisplaytoScreen" , "SendEmail" , "SendAttachment" , "Colour1" , "Colour2" , "TitleTxtColour" , "TimeToRun" , "PluginSeconds" , "Style" , "SQLExport", "Date")
foreach($vcvar in $vcvars) {
	if (!($(Get-Variable -Name "$vcvar" -erroraction 'silentlycontinue'))) {
		Write-Host "Variable `$$vcvar is not defined in GlobalVariables.ps1" -foregroundcolor "Red"
		Write-Host "Press any key to exit ..."
		$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		Exit
	} 
}


# If $SQLExport is required, check that Connection String file has been created and test connection to DB
If($SQLExport -eq $true)
    {
    $SqlSCripts = $ScriptPath + "\SaveToDatabase.ps1"
    . $SqlSCripts 
    $SQLStringPath = $ScriptPath + "\SqlConnectionString.bin"
    If (!(Test-Path $SQLStringPath -ErrorAction SilentlyContinue))
        {
        # Prompt for SQL Connection String and encrypt
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "Save To Database v1.0 by Alan van Wyk"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "==================================="
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "This is the first time DB backup has been requested"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "In order for succesful backup to DB, you'll need to provide an SQL connection string"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "The account writing to DB will require write access to the tables generated"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "If the tables do not exist, the account will need the rights to create the relevant tables"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "This connection string will be encrypted and only the account that creates the config" 
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "file will be able to decrypt it"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "In order to reset this, simply delete the file: $SQLStringPath"
        Write-Host -foreground $Warning_Foreground -background $Warning_Background
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "This will be stored (encrypted) at:" 
        Write-Host -foreground $Warning_Foreground -background $Warning_Background "For help with connection strings, see: http://www.connectionstrings.com/sql-server-2012/"
        $stringIn = read-host "Please a provide valid Connection String"
        Encrypt-String -string $stringIn -path $SQLStringPath
        }
    }



$StylePath = $ScriptPath + "\Styles\" + $Style
if(!(Test-Path ($StylePath))) {
	# The path is not valid
	# Use the default style
	Write-Debug "Style path ($($StylePath)) is not valid"
	$StylePath = $ScriptPath + "\Styles\Default"
	Write-Debug "Using $($StylePath)"
}

# Import the Style
. ("$($StylePath)\Style.ps1")


Function Get-Base64Image ($Path) {
	$pic = Get-Content $Path -Encoding Byte
	[Convert]::ToBase64String($pic)
}

Function Get-CustomHTML ($Header, $HeaderImg){
	$Report = $HTMLHeader -replace "_HEADER_", $Header
    $Report = $Report -replace "_HEADERIMG_", $HeaderImg
	Return $Report
}

Function Get-CustomHeader0 ($Title){
	$Report = $CustomHeader0 -replace "_TITLE_", $Title
	Return $Report
}

Function Get-CustomHeader ($Title, $Comments){
	$Report = $CustomHeaderStart -replace "_TITLE_", $Title
	If ($Comments) {
		$Report += $CustomheaderComments -replace "_COMMENTS_", $Comments
	}
	$Report += $CustomHeaderEnd
	Return $Report
}

Function Get-CustomHeaderClose{
	$Report = $CustomHeaderClose
	Return $Report
}

Function Get-CustomHeader0Close{
	$Report = $CustomHeader0Close
	Return $Report
}

Function Get-CustomHTMLClose{
	$Report = $CustomHTMLClose
	Return $Report
}

Function Get-HTMLTable {
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html -Fragment
	$HTMLTable = $HTMLTable -Replace '<TABLE>', $HTMLTableReplace
	$HTMLTable = $HTMLTable -Replace '<td>', $HTMLTdReplace
	$HTMLTable = $HTMLTable -Replace '<th>', $HTMLThReplace
	$HTMLTable = $HTMLTable -replace '&lt;', '<'
	$HTMLTable = $HTMLTable -replace '&gt;', '>'
	Return $HTMLTable
}

Function Get-HTMLDetail ($Heading, $Detail){
	$Report = ($HTMLDetail -replace "_Heading_", $Heading) -replace "_Detail_", $Detail
	Return $Report
}

# Adding all plugins
$TTRReport = @()
$MyReport = Get-CustomHTML "$Server vCheck"
$MyReport += Get-CustomHeader0 ($Server)
$Plugins | Foreach {
	$IDinfo = Get-PluginID $_.Fullname
	$Title = $IDinfo[0]
	$Ver = $IDinfo[1]
	$Author = $IDinfo[2]
	Write-CustomOut "..start calculating $Title by $Author v$Ver"
	$TTR = [math]::round((Measure-Command {$Details = . $_.FullName}).TotalSeconds, 2)
	$TTRTable = "" | Select Plugin, TimeToRun
	$TTRTable.Plugin = $_.Name
	$TTRTable.TimeToRun = $TTR
	$TTRReport += $TTRTable
	$ver = "{0:N1}" -f $PluginVersion
	Write-CustomOut "..finished calculating $Title by $Author v$Ver"

	If ($Details) 
    {
		If ($Display -eq "List"){
			$MyReport += Get-CustomHeader $Header $Comments
			$AllProperties = $Details | Get-Member -MemberType Properties
			$AllProperties | Foreach {
				$MyReport += Get-HTMLDetail ($_.Name) ($Details.($_.Name))
			}
			$MyReport += Get-CustomHeaderClose			
		}
		If ($Display -eq "Table") {
			$MyReport += Get-CustomHeader $Header $Comments
			$MyReport += Get-HTMLTable $Details
			$MyReport += Get-CustomHeaderClose
		}
	If($SQLExport -eq $true)
        {
        Write-CustomOut "..writing $Title to database by Alan van Wyk v1.0"
        Export-ToVCheckDB -InputObject $Details -ProviderType $ProviderName -Description $Title -InstanceName $Server -CreateTableIfDoesNotExist -ConnectionString (Get-EncryptedString -path $SQLStringPath) 
        }
    }
}	
$MyReport += Get-CustomHeader ("This report took " + [math]::round(((Get-Date) - $Date).TotalMinutes,2) + " minutes to run all checks.") "The following plugins took longer than $PluginSeconds seconds to run, there may be a way to optimize these or remove them if not needed"
$TTRReport = $TTRReport | Where { $_.TimeToRun -gt $PluginSeconds } | Sort-Object TimeToRun -Descending
$TTRReport |  Foreach {$MyReport += Get-HTMLDetail $_.Plugin $_.TimeToRun}
$MyReport += Get-CustomHeaderClose
$MyReport += Get-CustomHeader0Close
$MyReport += Get-CustomHTMLClose

if ($DisplayToScreen -or $SetupSetting) {
	Write-CustomOut "..Displaying HTML results"
	$Filename = $Env:TEMP + "\" + $Server + "vCheck" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + ".htm"
	$MyReport | out-file -encoding ASCII -filepath $Filename
	Invoke-Item $Filename
}

if ($SendAttachment) {
	$Filename = $Env:TEMP + "\" + $Server + "vCheck" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + ".htm"
	$MyReport | out-file -encoding ASCII -filepath $Filename
}

if ($Outputpath) {
	$DateHTML = Get-Date -Format "yyyyMMddHH"
	$ArchiveFilePath = $Outputpath + "\Archives\" + $VIServer
	if (-not (Test-Path -PathType Container $ArchiveFilePath)) { New-Item $ArchiveFilePath -type directory | Out-Null }
	$Filename = $ArchiveFilePath + "\" + $VIServer + "_vCheck_" + $DateHTML + ".htm"
	$MyReport | out-file -encoding ASCII -filepath $Filename
}

if ($SendEmail) {
	Write-CustomOut "..Sending Email"
	If ($SendAttachment) {
		send-Mailmessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -SmtpServer $SMTPSRV -Body "vCheck attached to this email" -Attachments $Filename
	} Else {
		send-Mailmessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -SmtpServer $SMTPSRV -Body $MyReport -BodyAsHtml
	}
}

if ($SendAttachment -eq $true -and $DisplaytoScreen -ne $true) {
	Write-CustomOut "..Removing temporary file"
	Remove-Item $Filename -Force
}

$End = $ScriptPath + "\EndScript.ps1"
. $End
