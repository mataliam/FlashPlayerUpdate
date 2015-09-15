<#
	.SYNOPSIS
		Get-FlashPlayerUpdate is a Powershell module to check and install FlashPlayer updates.

	.DESCRIPTION
		Recently and historically speaking, FlashPlayer has been one of the primary target of many cyber attacks. All users should keep the FlashPlayer up-to-date, a very important step to deter these threats.  There are 3 types of FlashPlayers on all Windows systems and to update any or all, the current process has been painful and bit confusing.  This user friendly module has been designed to check for updates and ease the updating of any or all types of FlashPlayers.
	
	.PARAMETER  Type
		There are 3 major types of FlashPlayers for Windows activex, plugin and pepper. This script provides additional [default] type all. As the name suggest 'all' checks/patches activex, plugin and pepper type of FlashPlayers.

	.PARAMETER  Patch
		Valid options for this parameter is yes, no. Default is yes. This parameter lets you check the version without patching the system.

	.PARAMETER  Logfile
		This is for future expansion of the module; where one can specify and logfile to redirect all the optput to that log file	

	.EXAMPLE
		PS C:\> Get-FlashPlayerUpdate -Type 'activex' -Patch 'no'
		Sample output:
		================================================================================
		Checking Flash Player version for activex architecture
		================================================================================
		Congratulations, your Flash Player (activex) [18.0.0.232] is up-to-date
		
		This example shows how to check Local as well as Latest versions of ActiveX type FlashPlayer (used in IE). The additional parameter -Patch 'no' instructs module that we just want to compare Local and Latest version without installing it.

	.EXAMPLE
		PS C:\> Get-FlashPlayerUpdate
		Sample output:
		================================================================================
		Checking Flash Player version for activex architecture
		================================================================================
		Congratulations, your Flash Player (activex) [18.0.0.232] is up-to-date
		================================================================================
		Checking Flash Player version for plugin architecture
		================================================================================
		No Plugin Flash Found
		Installing Flash for plugin
		You have successfully updated Flash Player for plugin [18.0.0.232]
		Total time taken[0 h:0 m:4 s]
		================================================================================
		Checking Flash Player version for pepper architecture
		================================================================================
		No Pepper Flash Found
		Installing Flash for pepper
		You have successfully updated Flash Player for pepper [18.0.0.232]
		Total time taken[0 h:0 m:5 s]
		
		This example takes advantage of default -Type parameter 'all'. In other words, this will go and check if there is an update available for activex, plugin or pepper type FlashPlayer; and if yes it will download it from Adobe and install it.

	.INPUTS
		System.String

	.OUTPUTS
		System.String

	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.

	.LINK
		about_modules
#>

function Get-FlashPlayerUpdate
{
	[CmdletBinding()]
	param (
		[ValidateSet('activex', 'plugin', 'pepper', 'all')]
		[string]$Type = 'all',
		[ValidateSet('yes', 'no')]
		[string]$Patch = 'yes',
		[string]$Logfile 
	)
	begin
	{
		# Verify and host version of Powershell is at least 3
		if ($PSVersionTable.PSVersion.Major -lt 3)
		{
			Write-Host "`r`nYou at least need PowerShell Version 3 to run some features of this script. Please install latest Windows Management Framework and re-run the script"
			Exit
		}
		try
		{
			$URI = 'http://www.adobe.com/software/flash/about/'
			# Creating a global variable HTML; 'Test-Path' tests if variable HTML already exist and if so don't repeat this operation which optimise performance of the module as it does not have to go fetch content from internet multiple times
			if (!(Test-Path variable:global:HTML))
			{
				# using standard URL by adobe where they publish their latest FlashPlayer versions, Invoke-WebRequest stores the content of whole page into variable HTML
				$global:HTML = Invoke-WebRequest -Uri $URI
				
				# A hack put together with trial and error extracts Version detail table. There is possibility of this being point of failiur but a RegEx verifies the version later in the module
				$global:tempvar = (($HTML.ParsedHtml.getElementsByTagName('table') | Where{ $_.className -eq 'data-bordered max' }).innerText).split("`n")
				
			}
		}
		catch
		{
			# if this operation did not work then there was a problem with internet and we can't go any further.
			Write-Debug "Some issue connecting it to Internet"
		}
	}
	process
	{
		####################################################################################################################
		#		 General notes: This function checks Local and Latest (Adobe) version of type activex, plugin and pepper
		#		 Input:
		#			Parameters (All mandatory))
		#				Location [Type String]: Local, adobe
		#				Type [Type String]: activex, plugin, pepper
		#		 Output:
		#			VersionNumber [Type String]
		####################################################################################################################
		
		function Get-FlashPlayerCurrVer
		{
			Param
			(
				[Parameter(Mandatory = $True)]
				[ValidateSet('local', 'adobe')]
				[string]$Location,
				[Parameter(Mandatory = $True)]
				[ValidateSet('activex', 'plugin', 'pepper')]
				[string]$Type
			)
			
			Switch ($Location)
			{
				# Local versions are extracted from Local Machine registry hive
				"local" {
					Switch ($Type)
					{
						"activex" {
							if (Test-Path hklm:\SOFTWARE\Macromedia\FlashPlayerActiveX)
							{
								$VersionNumber = (Get-ItemProperty hklm:\SOFTWARE\Macromedia\FlashPlayerActiveX).Version
							}
							else
							{
								# logic infers that if no registry found, Flash player for that type does not exist
								$VersionNumber = "No ActiveX Flash Found"
							}
						}
						"plugin" {
							if (Test-Path hklm:\SOFTWARE\Macromedia\FlashPlayerPlugin)
							{
								$VersionNumber = (Get-ItemProperty hklm:\SOFTWARE\Macromedia\FlashPlayerPlugin).Version
							}
							else
							{
								$VersionNumber = "No Plugin Flash Found"
							}
						}
						"pepper" {
							if (Test-Path hklm:\SOFTWARE\Macromedia\FlashPlayerPepper)
							{
								$VersionNumber = (Get-ItemProperty hklm:\SOFTWARE\Macromedia\FlashPlayerPepper).Version
							}
							else
							{
								$VersionNumber = "No Pepper Flash Found"
							}
						}
					}
					
				}
				"adobe" {
			
					# use the extracted info from the tempvar array. From trial and error we know at location 1,3,4 we can find version number entry for activex, plugin and pepper architecture respectively
					# version number is extracted by removing everything including these words (ActiveX,NPAPI,PPAPI) from left of that string
					Switch ($Type)
					{
						"activex" {
							$CurVerActiveX = $tempvar[1]
							$tempIndex = $CurVerActiveX.IndexOf('ActiveX') + 7
							$VersionNumber = $CurVerActiveX.Substring($tempIndex, $($CurVerActiveX.Length - $tempIndex))
						}
						"plugin" {
							$CurVerPlugin = $tempvar[3]
							$tempIndex = $CurVerPlugin.IndexOf('NPAPI') + 5
							$VersionNumber = $CurVerPlugin.Substring($tempIndex, $($CurVerPlugin.Length - $tempIndex))
						}
						"pepper" {
							$CurVerPepper = $tempvar[4]
							$tempIndex = $CurVerPepper.IndexOf('PPAPI') + 5
							$VersionNumber = $CurVerPepper.Substring($tempIndex, $($CurVerPepper.Length - $tempIndex))
						}
					}
					
					# finally if for what ever reasons that Adobe site request had not worked, or the extracted version number of FlashPlayer is not in the format of [X.X.X.X] where X is any digit; then an internet explorer window is opened with 
					# URL which has link to download etc.
					if (($HTML.StatusCode -ne 200) -or !($VersionNumber -match '\d+.\d+.\d+.\d+'))
					{
						Write-Host "`r`n`r`nSorry there was some issue with getting Adobe version. A window would shortly open where you can refer to it manually."
						Start-Process "C:\Program Files\Internet Explorer\iexplore.exe" -Args $URI
						exit
					}
				}
			}
			return $VersionNumber.Trim()
		}
		
		####################################################################################################################
		#		 General notes: This function obtains Latest (Adobe) version of type activex, plugin or pepper and installs it
		#		 Input:
		#			Parameters (All mandatory))
		#				Type [Type String]: activex, plugin, pepper
		#		 Output:
		#			VersionNumber [Type String]
		####################################################################################################################
		
		function Install-FlashPlayer
		{
			param
			(
				[Parameter(Mandatory = $True)]
				[ValidateSet('activex', 'plugin', 'pepper')]
				[string]$Type
			)
			# A temp directory in C root is created if not already exist to store the installer
			if (!(Test-Path "c:\Utilities"))
			{
				New-Item C:\Utilities -ItemType directory
				# this global variable helps cleanup after the fact if that folder did not exist before
				$global:TempFolder = $true
			}
			
			# A filename with path is generated on the fly based on the function parameter 'type'
			$Local_storage = "C:\Utilities\FP_$($Type)_installer.exe"
			switch ($Type)
			{
				'activex' { $Uvar = '_ax' }
				'plugin' { $Uvar = '' }
				'pepper' { $Uvar = '_ppapi' }
			}
			
			# Adobe provides latest FlashPlayer with some suffix at the end, which we create with above switch and incorporate it in bellow URL
			$Adobe_url = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player$Uvar.exe"
			
			# There are a few ways to download a file from web using PowerShell but this .net method proven to be most efficient. This is available all the way since .net 1.1
			(New-Object System.Net.WebClient).DownloadFile($Adobe_url, $Local_storage)
			
			# To test the install Global $error variable is cleared and LASTEXITCODE is set to NULL 
			$error.clear()
			$global:LASTEXITCODE = $null
			Write-Debug "Installing $Local_storage..."
			
			# A handy command in PowerShell allows measuring of time taken for this install process
			# Installer exe takes command argument '-install' to perform a silent install (which can be found in Admin guide for FlashPlayer)
			$TotalTime = Measure-Command { Start-Process $Local_storage -Args '-install' -Wait }
			
			# bellow if statement checks for any error reported and the exit code be anything else than 0. 
			if (($? -eq $true) -and ([string]::IsNullOrEmpty($error[0])) -and ([string]::IsNullOrEmpty($lastexitcode) -or $lastexitcode -eq 0))
			{
				$Local_Version = Get-FlashPlayerCurrVer -Location 'local' -Type $Type
				$Adobe_Version = Get-FlashPlayerCurrVer -Location 'adobe' -Type $Type
				
				# If all has gone well, no errors and exit code is also 0, one more time version is checked and if for what ever reasons its not same as latest version
				# (Eg: Starting Windows 8.1, Windows does not allow install of ActiveX based FlashPlayer, as they have embeded the FlashPlayer in IE) then a message is shown on the screen suggesting same
				# if version matches, a message is also shown on the screen with the version of it and total time it took to get installed
				if ($Local_Version -eq $Adobe_Version)
				{
					Write-Host "`r`nYou have successfully updated Flash Player for $Type [$Local_Version]`r`nTotal time taken[$($TotalTime.Hours) h:$($TotalTime.Minutes) m:$($TotalTime.Seconds) s]"
				}
				else
				{
					Write-Host "`r`nSomething went wrong. Adobe's current version for $Type is $Adobe_Version; and local version of $Type is $Local_Version"
				}
			}
			else
			{
				# there are some listed exit codes by Adobe which shows proper error
				switch ($LASTEXITCODE)
				{
					1003 { $ErrorMessage = "Invalid argument passed to installer" }
					1011 { $ErrorMessage = "Install already in progress" }
					1012 { $ErrorMessage = "Does not have admin permissions (W2K, XP)" }
					1013 { $ErrorMessage = "Trying to install older revision" }
					1022 { $ErrorMessage = "Does not have admin permissions (Vista, Windows 7)" }
					1024 { $ErrorMessage = "Unable to write files to directory" }
					1025 { $ErrorMessage = "Existing Player in use" }
					1032 { $ErrorMessage = "ActiveX registration failed" }
					1041 { $ErrorMessage = "An application that uses the Flash Player is open. Quit the application and try again." }
					3 { $ErrorMessage = "Does not have admin permissions" }
					4 { $ErrorMessage = "Unsupported OS" }
					5 { $ErrorMessage = "Previously installed with elevated permissions" }
					6 { $ErrorMessage = "Insufficient disk space" }
					7 { $ErrorMessage = "Trying to install older revision" }
					8 { $ErrorMessage = "Browser is open" }
				}
				Write-Host "`r`nThere was en error with update. Specific error was:`r`n$ErrorMessage"
			}
			
			# At the end of install success or not, temp file is deleted
			Remove-Item $Local_storage -Force
			return $Result
		}
		
		# TODO: in future implement this function to redirect output of the script to a logfile if specified
		function Log-this
		{
			param
			(
				[parameter(Mandatory = $true)]
				[string]
				$message
			)
			if ($Logfile.Trim() -eq '')
			{
				Write-Host $message
			}
			else
			{
				Write-Output $message | Out-File $Logfile -Append
			}
		}
		
		if ($Logfile.Trim() -ne '')
		{
			if (!(Test-Path $Logfile))
			{
				New-Item  -ItemType file -Path $Logfile
			}
		}
		
		# if the type is 'all' then this code, calls the mother function Get-FlashPlayerUpdate with appropriate parameter each time for all 3 types
		if ($Type -eq 'all')
		{
			$CommandToRun = ("Get-FlashPlayerUpdate -Type 'activex' -Patch $Patch")
			Write-Debug "Running:`r`n$CommandToRun"
			Invoke-Expression $CommandToRun
			$CommandToRun = ("Get-FlashPlayerUpdate -Type 'plugin' -Patch $Patch")
			Write-Debug "Running:`r`n$CommandToRun"
			Invoke-Expression $CommandToRun
			$CommandToRun = ("Get-FlashPlayerUpdate -Type 'pepper' -Patch $Patch")
			Write-Debug "Running:`r`n$CommandToRun"
			Invoke-Expression $CommandToRun
		}
		else
		{
			# if a type is specified during fuction call, a message is showen on the screen
			Write-Host $("=" * 80)
			Write-Host "Checking Flash Player version for $Type architecture"
			Write-Host $("=" * 80)
			
			# Local version of the respective FlashPlayer type is obtian via our function call
			$Local_Version = Get-FlashPlayerCurrVer -Location 'local' -Type $Type
			
			# If there is no FlashPlayer detected, our response will start from 'No' which is by design to show details with in same response then if portion is executed
			if ($Local_Version.Substring(0, 2) -eq "No")
			{
				# if parameter Patch is set to yes then Install-FlashPlayer function is called else a message is shown on the screen suggesting they should install latest version
				Write-Host "`r`n$Local_Version`r`n"
				if ($Patch -eq 'yes')
				{
					Write-Host "`r`nInstalling Flash for $Type"
					Install-FlashPlayer -Type $Type
					Write-Host "`r`n"
				}
				else
				{
					Write-Host "`r`nPlse install Flash Player ($Type), your system is at very high risk`r`n`r`n"
				}
			}
			else
			{
				# We will be here if the script has found a version number from registry hive
				
				# We get Version from Adobe
				$Adobe_Version = Get-FlashPlayerCurrVer -Location 'adobe' -Type $Type
				
				# if both local and aodbe version are same message is shown on the screen suggesting you are up-to-date
				if ($Local_Version -eq $Adobe_Version)
				{
					Write-Host "`r`nCongratulations, your Flash Player ($Type) [$Adobe_Version] is up-to-date`r`n`r`n"
				}
				else
				{
					# we will come here if local version and adobe version don't match. We display appropriate message on the screen
					
					Write-Host "`r`nCurrent version of Flash Player is: $Adobe_Version, version on your computer is: $Local_Version`r`n"
					
					# if patching is allowed we call function Install-FlashPlayer and supply type to install FlashPlayer
					if ($Patch -eq 'yes')
					{
						Write-Host "`r`nUpdating your Flash Player...`r`n"
						Install-FlashPlayer -Type $Type
					}
					else
					{
						# else we let user know that they should install latest FlashPlayer
						Write-Host "`r`nPlse update your Flash Player ($Type), your system is at very high risk`r`n`r`n"
					}
				}
			}
		}
	}
	end
	{
		# At the end we delete temp folder it was not there
		try
		{
			if ($TempFolder)
			{
				Remove-Item -Path C:\Utilities -Force -Recurse >$null 2>&1
			}
		}
		catch
		{
			Write-Debug "Something went wrong while removing Utilities folder"
		}
	}
}
Export-ModuleMember -Function Get-FlashPlayerUpdate

# Alias is created to assist quick function
New-Alias -Name gfpu -Value Get-FlashPlayerUpdate
Export-ModuleMember -Alias gfpu