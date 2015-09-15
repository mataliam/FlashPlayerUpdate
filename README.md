# PowerShell module to update FlashPlayer
<h3>Background</h3>

Recently and historically speaking, FlashPlayer has been one of the primary target of many cyber attacks. All users should keep the FlashPlayer up-to-date, a very important step to deter these threats.  There are 3 types of FlashPlayers on all Windows systems and to update any or all, the current process has been painful and bit confusing for average user. Like everyone I felt the same pain so I wrote this simple PowerShell module which can check for updates and make the process easier and efficient.

<h3>Pre-requisite</h3>

To use this script you would need PowerShell version 3 or higher. General rule of thumb is, if you have Windows 8 or higher then you are good. If you are still on Windows 7, get latest Windows management framework. As of this writing, most recent Windows Management Framework (WMF) is 4. First make sure you have latest .Net Framework [as of writing this: Web-link] and then get the WMF [Web-link]

<h3>Step 1: import the module</h3>

Download FlashPlayerUpdate.zip, extract the file and copy ‘FlashPlayerUpdate’ folder to:
'C:\Windows\System32\WindowsPowerShell\v1.0\Modules'
Be careful that files are not in a sub folder of FlashPlayerUpdate. Alternatively if you prefer not to import this module to above path, refer this [web-link]
Run PowerShell as Administrator (Right click on PowerShell & choose Run As Administrator)
Make sure execution policy is set to Remotesigned [run command Set-ExecutionPolicy RemoteSigned]
Run command: Import-Module FlashPlayerUpdate

<h3>Step 2: Use it :-)</h3>

Run command Get-FlashPlayerUpdate
Alternatively you can also run this script by running command gfpu.
Get-Help Get-FlashPlayerUpdate -Full will give you more details on how this can be used.
