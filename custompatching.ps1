﻿######################### Notes and Explanations #########################
#Variables
<#
    .variable $ErrorActionPreference
    Type = NONE
    https://goo.gl/T6LA2J
---------------------------------------------------------------------------------
    .variable $AutomaticVariables
     Type = Object
     This variable is a list of every existing variable at the start of the script
     run. It gives us a comparison point later when you want to get the values of
     all variables that the script has created.
---------------------------------------------------------------------------------
    .variable $ErrorPath
    Type = String
    When the script ends this path will be used to output the contents of $Error.
    This is for script troubleshooting purposes.
---------------------------------------------------------------------------------
    .variable $SQLInsertPath
    This variable contains a filepath used to store the SQLInserts we pull back
    into LabTech.
---------------------------------------------------------------------------------
    .variable $Filter
    The variable used to prune out patches. By default i am filtering patches that:
        1) are Installed
        2) are Hidden
        3) Require a reboot

    Additional filters can be applied via the $WantedUpdateTypeIDs variable 
    and the $strExcludedKBs variable.
---------------------------------------------------------------------------------
    .variable $LogFilePath
    Type = String
    This is the path used by the Write-Log function.
---------------------------------------------------------------------------------
    .variable $OutVarPath
    Type = String
    When the script ends this path will be used to output the values of all user
    created variables. This is for script troubleshooting purposes.
---------------------------------------------------------------------------------
    .variable $ResultsPath
    Type = String
    When the script ends this path will be used to output the result of the script.
    LabTech then reads the contents of that file to know how to continue.
---------------------------------------------------------------------------------
    .variable $ScriptName
    Type = String
    The name of the script. Mostly used for logging purposes.
---------------------------------------------------------------------------------
    .variable $Computerid
    Type = String
    This variable gets autofilled from LabTech when the file is written down to the
    local machine. It is a unique identifier of the machine the script is being run
    against. This variable is only used in the making of the MySQL insert statements.
---------------------------------------------------------------------------------
    .variable $strExcludedKBs
    Type = String
    Pulled in from a LabTech extradatafield. Needs to be a comma delimited list of 
    KBs you DONT want installed. Used for filtering purposes. Works if null.
    Example: "KB8675309,KB1234567"
---------------------------------------------------------------------------------
    .variable $strWantedCategoryIDs
    Type = String
    Pulled in from LabTech. These are guids that correspond to certain patch categories. 
    See below for a full list of categories.
    Gathered From : https://goo.gl/jk4lzf

    Possible Category ID's the string can contain:

    Application                5C9376AB-8CE6-464A-B136-22113DD69801
    Connectors                 434DE588-ED14-48F5-8EED-A15E09A991F6
    CriticalUpdates            E6CF1350-C01B-414D-A61F-263D14D133B4
    DefinitionUpdates          E0789628-CE08-4437-BE74-2495B842F43B
    DeveloperKits              E140075D-8433-45C3-AD87-E72345B36078
    FeaturePacks               B54E7D24-7ADD-428F-8B75-90A396FA584F
    Guidance                   9511D615-35B2-47BB-927F-F73D8E9260BB
    SecurityUpdates            0FA1201D-4330-4FA8-8AE9-B877473B6441
    ServicePacks               68C5B0A3-D1A6-4553-AE49-01D3A7827828
    Tools                      B4832BD8-E735-4761-8DAF-37F882276DAB
    UpdateRollups              28BC880E-0592-4CBF-8F95-C79B17911D5F
    Updates                    CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD83
---------------------------------------------------------------------------------
    .variable $arrExcludedKBs
    Type = Array
    Contains all of the KB's in $strExcludedKBs in an array format. I 
    had to do this because the @variablenamehere@ way that Labtech variables 
    are named throws an error in an array variable format. It thinks it 
    is a splatting operator. So by making the string variable first, then
    putting that in the array it fixes this.
---------------------------------------------------------------------------------
    .variable $WantedUpdateTypeIDs
    Type = Array
    The wanted update types. All possible IDs are defined in the Set-UpdateTypeID function.
---------------------------------------------------------------------------------
#>

#Sections
<#
    .section Function Declarations
    Where all script functions are declared.
    Functions are listed alphabetically.
---------------------------------------------------------------------------------
    .section Variable Declarations
    Where possible, all variables are declared here. Full explanation of
    the variables is available above in the Variable Explanations Section.
---------------------------------------------------------------------------------
    .section Pre-Patch Checks
    Checks to see how much freespace is remaing on the OS drive. Also
    checks to see if there is a pending reboot which will break patching.
    If so returns Reboot Needed to LabTech. LabTech can then perform a reboot
    and rerun the script.
---------------------------------------------------------------------------------
    .section Define Pre Search Filters
    Checks to see if other filters were set such as $strWantedCategoryIDs.
    If they are it updates the filter variable to include those.
    ***NOT CURRENTLY FUNCTIONAL***
---------------------------------------------------------------------------------
    .section Get All Available Updates
    Performs the actual call to Microsoft/Windows Update and gathers the
    available patches. It also writes to the log the total number it gathered and
    how long the query took. It checks to verify if there are any needed patches and
	if there aren't it returns "No Patches Needed".
---------------------------------------------------------------------------------
    .section Download all the Updates
    This section pre-downloads all the updates to speed up the installation process.
---------------------------------------------------------------------------------
    .section Install the desired updates
    Filters out the updates we dont want and for each one we do, runs the
    Process-Update  and Determine Result functions against it. It then records 
    the success or failure of the patch in the update object.
---------------------------------------------------------------------------------
    .section Determine Exit Code
    Determines whether or not a reboot is needed and out-files the result.
---------------------------------------------------------------------------------
#>

#Possible Exit Values
<#
    Reboot Needed
    This value is returned when the script determines that a reboot is
    required before patching can be successfully started.

    No Patches Needed
    This value is returned when $searchresults.update.count equals 0. The
    script completed successfully but no patches needed to be installed.

    Success
    The script completed successfully and no reboot is required afterwards.

    Success -r
    The script completed successfully but the machine requires a reboot
    before it should be taken out of maintenance mode.

	Path Failure
	The script failed. The $SqlInsertPath variable did not resolve to a location we could reach.

#>

##########################################################################
#Function Declarations

Function SendTo-LogEntries
{
    Param
    (
		[Parameter(Mandatory = $true,Position = 0)]
		[STRING]$Token,
		[Parameter(Mandatory = $true,Position = 0)]
		[STRING]$Message   
    )
    $tcpConnection = New-Object System.Net.Sockets.TcpClient('data.logentries.com', '80')
    $tcpStream = $tcpConnection.GetStream()
    $reader = New-Object System.IO.StreamReader($tcpStream)
    $writer = New-Object System.IO.StreamWriter($tcpStream)
    $writer.AutoFlush = $true
    $buffer = new-object System.Byte[] 1024
    $encoding = new-object System.Text.AsciiEncoding 
    $writer.WriteLine("$Token $Message")
    write-output "$Token $Message"
    $reader.Close()
    $writer.Close()
    $tcpConnection.Close()
}

Function Create-InsertStatement
{
    <#
	.SYNOPSIS
		Converts a PowerShell Object to a SQLInsert Statement.
	
	.DESCRIPTION
		An Object is passed and each property of the object is used as a MySQL
        column name. The value for that objects property is used as the value.
	
	.PARAMETER Update
		The update object that contains all the properties to use in the insert.
	
	.PARAMETER Results
		The object that contains the results of the download and installation
		of the update.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Create-InsertStatement -Update $Update -ParsedInstallResult $ParsedInstallResult
    #>
	
	[CmdLetBinding()]
	Param
		(
		[Parameter(Mandatory = $False)]
		[Object]$Update,
		[Parameter(Mandatory = $False)]
		[Object]$Results
	)

    $kbid = $Update.KBArticleIDs | out-string
    $Title = $Update.title
    $Cat = $Update.categories.Item(0).name
    $Desc = $Update.Description
    $Guid = $Update.Identity.UpdateID
    $Sev = $Update.MSRCSeverity
    $Res = $Results.installresult
    $Hres = $Results.Hresultcode
    $HresDesc = $Results.Hresultdescription


    $Columns = @"
Insert into `Plugin_Patching_Data`
(`ComputerID`,
`KB_Number`,
`KB_Title`,
`KB_Category`,
`KB_Description`,
`KB_Guid`,
`Severity`,
`Result`,
`hResult`,
`hResultDesc`,
`DateInstalled`) VALUES 
({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10});
"@
 
    $MySQLInsert = $Columns -f $computerID, $kbid, $Title, $Cat, $Desc, $Guid, $Sev, $Res, $Hres, $HresDesc, 'NOW()'

    Add-Content -Path $SqlInsertsPath -Value $MySQLInsert
	SendTo-LogEntries -Token $Token -Message "Insert Statement for KB $($Kbid) added."
	
	If(!(Test-Path -Path $SqlInsertsPath))
	{
		SendTo-LogEntries -Token $Token -Message "SQL Insert Filepath is innacessible. Exiting script."
		End-Script -Result "Path Failure"
	}
}

function Determine-Result
{
    <#
	.SYNOPSIS
		Determines if an update installed successfully.
	
	.DESCRIPTION
        This function will determine the success or failure of an update
        installation and create a new object to reflect the results.
	
	.PARAMETER $InstallResult
        The result object passed back from windows update.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Determine-Result -InstallResult $Installresult

    #>
	
	param
		(
		[Parameter(Mandatory = $true, Position = 0)]
		$InstallResult
		)
	
	$InstallResultCode = $InstallResult.resultcode
	$InstallResultHref = $InstallResult.hresult
	[Array]$objInstallResults = @()
	
	If ($InstallResultCode -ne 2)
	{
		$objInstallResults += New-Object PSObject -Property @{
			InstallResult = "Failed to Install";
			HresultCode = $InstallResultHref;
			HresultDescription = Get-HresultDescription $InstallResultHref;
		}
		
	}
	
	Else
	{
		$objInstallResults += New-Object PSObject -Property @{
			InstallResult = "Success";
			HresultCode = $InstallResultHref;
			HresultDescription = "Success";
		}
	}

    Return $objInstallResults
}

Function End-Script
{
	
	<#
	.SYNOPSIS
		A function to wrap up the end of the script.
	
	.DESCRIPTION
		Function has multiple tasks:
		1) Out-files a list of user created variables and their values.
		2) Out-files the contents of $Error.
		3) Out-files $Result
		4) Terminates the Script.
	
	.PARAMETER $Result
		The result string to outfile.
	
	.EXAMPLE
		PS C:\> End-Script -Result $Result
	
	.NOTES
		N/A
	#>
	
	param
		(
		[parameter(Mandatory = $true)]
		[String]$Result
	)
	Out-File -InputObject $Error -FilePath $ErrorPath
	Out-File -InputObject $Result -Filepath $ResultsPath
    SendTo-LogEntries -Token $Token -Message"***** $($ScriptName) Ends *****"
	exit;
}

Function Format-SanitizedString
{
    	<#
	.SYNOPSIS
		A function to sanitize a string for insert into MySQL.
	
	.DESCRIPTION
		This function replaces special characters that dont play nice
        with MySQL Inserts.
	
	.PARAMETER $InputString
		The string to sanitize.
	
	.EXAMPLE
		PS C:\> Format-SanitizedString -InputString $String
	
	.NOTES
		N/A
	#>
	
	[CmdLetBinding()]
	Param
		(
		[Parameter(Mandatory = $False)]
		[String]$InputString
	)
	
	$SanitizedString = "";
	If ($InputString -ne $null -and $InputString.Trim().Length -gt 0)
	{
		$SanitizedString = $InputString.Trim();
		$SanitizedString = $SanitizedString.Replace("\", "\\");
		$SanitizedString = $SanitizedString.Replace("'", "\'");
		$SanitizedString = $SanitizedString.Replace("`"", "\`"");
	}
	
	Return $SanitizedString
}

function Get-HresultDescription
{
    <#
	.SYNOPSIS
		Attempts to get a detailed reason for why an update failed to 
        download or install.
	
	.DESCRIPTION
		This function queries a MSDN page that lists the hresult codes 
        and pulls back the listed description for that code.
	
	.PARAMETER HresultCode
		The Microsofthresult code of the update.
        Ex: "0x80240044"
	
	.NOTES
		N/A
	
	.EXAMPLE
		Get-HresultDescription '0x80240044'

    .EXAMPLE
        Get-HresultDescription $HResult
#>
	
	param
		(
		[Parameter(Mandatory = $true, Position = 0)]
		[String]$HresultCode
	)
	
	$Webresult = (Invoke-WebRequest -URI https://technet.microsoft.com/en-us/library/cc720442.aspx).content
	$regex = "(?:<p>$HresultCode<\/p>[\s]+<\/td>[\s]+<td.+\s+(?:.+\s+){3}<p>(.+))(?:</p>)"
	[String]$Hresultcode = ([regex]::matches($Webresult, $regex)).groups[1].value
	
	Return [String]$Hresultcode
}

function Process-Update
{
    <#
	.SYNOPSIS
		Downloads and installs a Windows Update
	
	.DESCRIPTION
		This function downloads a specific windows update to prepare
        for installation.Then the update is installed and hresult code is
        checked for failures.
	
	.PARAMETER UpdateID
		The Microsoft ID of the update. (NOT THE KB NUMBER)
        Ex: "f1b1a591-bb75-4b1c-9fbd-03eedb00cc9d"
	
	.NOTES
		N/A
	
	.EXAMPLE
		Process-Update 'f1b1a591-bb75-4b1c-9fbd-03eedb00cc9d'

    .EXAMPLE
        Process-Update $Id
#>
	
	param
		(
		[Parameter(Mandatory = $true, Position = 0)]
		$Update
	)
	
	$TempCollection = New-Object -ComObject Microsoft.Update.UpdateColl
	$tempcollection.add($Update)
	$TempInstaller = New-Object -ComObject Microsoft.Update.Installer
	$TempInstaller.Updates = $TempCollection
	$ProcessResult = $TempInstaller.Install()
	
	Return $ProcessResult
}

function Set-UpdateTypeID
{
    <#
	.SYNOPSIS
		Sets a Windows Update Type ID
	
	.DESCRIPTION
		The Descriptive term for an update is passed and a numerical representation is returned.
	
	.PARAMETER UpdateType
		The Descriptive term of the update category.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Set-UpdateTypeID 'Critical'
    #>
	
	param
		(
		[Parameter(Mandatory = $true, Position = 0)]
		[String]$UpdateType
	)
	switch -wildcard ([String]$UpdateType)
	{
		"Critical*"             { [INT]$UpdateID = 0 }
		"Definition Updates"    { [INT]$UpdateID = 1 }
		"Drivers"               { [INT]$UpdateID = 2 }
		"Feature Packs"         { [INT]$UpdateID = 3 }
		"Security*"             { [INT]$UpdateID = 4 }
		"ServicePacks"          { [INT]$UpdateID = 5 }
		"Tools"                 { [INT]$UpdateID = 6 }
		"*Rollup*"              { [INT]$UpdateID = 7 }
		"Updates"               { [INT]$UpdateID = 8 }
		"Microsoft"             { [INT]$UpdateID = 9 }
		default { [INT]$UpdateID = 99 }
	}
	
	Return [INT]$UpdateID
}

##########################################################################
#Variable Declarations

$ErrorActionPreference = 'SilentlyContinue'
[String]$Computerid = "@ComputerID@"
[String]$Filter = 'IsInstalled = 0 and IsHidden=0'
[String]$ErrorPath = "$($env:windir)\temp\PatchingAutomationERRORS.txt"
[String]$ResultsPath = "$($env:windir)\temp\PatchingAutomationRESULTS.txt"
[String]$ScriptName = 'Custom Patching'
[String]$SqlInsertsPath = "$($env:windir)\temp\sqlinserts.txt"
#[String]$strExcludedKBs = "@ExcludedKBs@"
[String]$strExcludedKBs = ""
#[String]$strWantedCategoryIDs = '@PatchCategories@'
[String]$strWantedCategoryIDs = ""
[Array]$arrExcludedKBs = @($strExcludedKBs)
[Array]$WantedUpdateTypeIDs = @(0, 4, 5, 7, 8, 9)
[String]$Token = 'e119e037-16d0-4a5f-aa70-ded70a6682e5'

##########################################################################
#Pre-Patch Checks

SendTo-LogEntries -Token $Token -Message "**** $($ScriptName) Begins ****"
SendTo-LogEntries -Token $Token -Message "Computer: $env:COMPUTERNAME"
SendTo-LogEntries -Token $Token -Message "OS:       $((Get-WmiObject Win32_OperatingSystem).caption)"

IF (Test-Path $ErrorPath -ErrorAction SilentlyContinue)		 {Remove-Item $ErrorPath -Force}
IF (Test-Path $OutVarPath -ErrorAction SilentlyContinue) 	 {Remove-Item $OutVarPath -Force}
IF (Test-Path $ResultsPath -ErrorAction SilentlyContinue) 	 {Remove-Item $ResultsPath -Force}
IF (Test-Path $SqlInsertsPath -ErrorAction SilentlyContinue) {Remove-Item $SqlInsertsPath -Force}

$FreeSpace = (Get-WmiObject -class win32_LogicalDisk -filter "Name = 'C:'").freespace
$FreespaceMB = "{0:n2}" -f ($Freespace/1MB)
SendTo-LogEntries -Token $Token -Message "Freespace before patching (in Megabytes) : $FreeSpaceMB"

##########################################################################
#Define Pre Search Filters

if ($strwantedcategoryids)
{
	$filter = $filter + " and CategoryIDs contains " + $strWantedCategoryIDs
}

SendTo-LogEntries -Token $Token -Message "Pre-Patch filter is : $filter"

##########################################################################
#Get all available updates

SendTo-LogEntries -Token $Token -Message "Beginning Search to Gather Updates."
$Start = Get-Date
$objSession = New-Object -com "Microsoft.Update.Session"
$objSearcher = $objSession.CreateUpdateSearcher()
$serviceName = "Windows Update"
$SearchResults = $objSearcher.Search("$filter")
$End = Get-Date
$SearchTimespan = new-timespan -Start $start -End $end
SendTo-LogEntries -Token $Token -Message "Update Search Completed. Time taken was $($Searchtimespan.totalseconds) seconds"
SendTo-LogEntries -Token $Token -Message "There are $($SearchResults.Updates.Count) Total updates available."

If ($SearchResults.Updates.Count -eq 0)
{
	SendTo-LogEntries -Token $Token -Message "Script has determined that no patches are required."
	End-Script -Result "No Patches Needed"
}

##########################################################################
#Download all the Updates

SendTo-LogEntries -Token $Token -Message "Beginning download of all updates..."
$Start = Get-Date
$Downloader = $objSession.CreateUpdateDownloader()
$Downloader.updates = $SearchResults.updates
$DownloadResults = $Downloader.Download()
$End = Get-Date
$DownloadTimespan = new-timespan -Start $start -End $end
SendTo-LogEntries -Token $Token -Message "Downloads Completed.Time taken $Downloadtimespan.totalseconds"

If ($DownloadResults.resultcode -ne 2)
{
    SendTo-LogEntries -Token $Token -Message "Download Results were: Downloading failed for one or more patches."
}

Else
{
    SendTo-LogEntries -Token $Token -Message "Download Results were: All Patches Downloaded Successfully."
}


##########################################################################
#Install the desired updates

[INT]$Count = 1

foreach ($Update in $Searchresults.updates)
{

   SendTo-LogEntries -Token $Token -Message "Processing Update $Count of $($SearchResults.Updates.Count)" 

	If ($arrExcludedKBs -contains $update.KB)
	{
		SendTo-LogEntries -Token $Token -Message "KB $($Update.kb) was excluded."
	}
	
	Else
	{
		SendTo-LogEntries -Token $Token -Message "KB ID: $($Update.kbarticleids)"
        SendTo-LogEntries -Token $Token -Message "KB Title: $($Update.title)"
        SendTo-LogEntries -Token $Token -Message "KB Type: $($Update.type)"
		$InstallResult = Process-Update $Update
		$ParsedInstallResult = Determine-Result -InstallResult $Installresult
		SendTo-LogEntries -Token $Token -Message "Result was: $($ParsedInstallResult.installresult)"
		
		foreach ($Property in $Update.psobject.properties)
		{
			$Property.Value = Format-SanitizedString $Property.Value
		}
	}

    [INT]$Count++
}

##########################################################################
#Determine Exit Code

If ($objSystemInfo.RebootRequired -eq $True)
{
	SendTo-LogEntries -Token $Token -Message "Script completed successfully but a reboot is required."
	End-Script -Result "Success -r"
}

Else
{
	SendTo-LogEntries -Token $Token -Message "Script completed successfully and a reboot is not required."
	End-Script -Result "Success"
}
