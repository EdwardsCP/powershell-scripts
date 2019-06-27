<#
VaronisDisableUserDenySMB.ps1
Original Script provided by Varonis "Disable-user-Deny-SMB.6.2v3.0.ps1"
Modified by Colin Edwards
June 2019
Revisions: Removed sections designed for 2008 and older File Servers.  Added code to generate a powershell script that can be run to remove the denies added to SMB shares, and enable to AD account.
-----------------------------

This script will disable the user account and set a deny on the SMB shares on the file server.
Setting a deny on SMB shares is required for an adequate rapid response to potential ransomware.  Disabling a user account does not trigger Urgent Replication.  It takes up to 15 minutes for a Disable action to replicate across all DCs.  So, for rapid response, we need to add a deny on the file share.


Requirements:
- Install Remote Server Administration Tools on the Varonis IDU and Varonis Collectors
- Enable PowerShell remoting on all monitored file servers - https://technet.microsoft.com/en-us/library/hh849694.aspx
- Execute Powershell using the "sysnative" path - C:\Windows\sysnative\WindowsPowerShell\v1.0\powershell.exe
- Only works on Windows Server 2008 or later File Servers
#>

#Start Transcript logging
Start-Transcript -Path 'D:\Powershell Transcripts\transcript.log' -Append

#Set Local Variables
$FileServer = $env:FileServerDomain
$ActingUser = $env:ActingObjectSAMAccountName
$ActingDomain = $env:ActingObjectDomain

#For testing, uncomment variables below and comment out variables above
#$FileServer = "FileServer"
#$ActingUser = "Test_User"
#$ActingDomain = "Domain_Name"

#Import required modules
Import-Module ActiveDirectory

#Disable the user's AD account
Disable-ADAccount $ActingUser

#Get the OS Version of the File Server
$OSVersion = Get-WmiObject -Computer $FileServer -Class Win32_OperatingSystem | Select -expand Caption

#Block SMB access on all shares on the file server
Invoke-Command -ComputerName $FileServer -ScriptBlock { ForEach ($Share in Get-SmbShare) { Block-SmbShareAccess -Name $Share.Name -AccountName $args[0] -Force } } -Args $ActingUser

#Create a PowerShell script that can be  used to unblock SMB access on all shares on the file server after the incident is investigated and resolved
$UnblockScriptName = "D:\PowerShell Unblock SMB Shares Scripts\" + $ActingUser + $FileServer + ".ps1"
Write-Output "<#" | Out-File -FilePath "$UnblockScriptName" -Append
Write-Output "Script Automatically Generated to Enable AD Account and undo an SMB Block created by VaronisDisableUserDenySMB1.ps1" | Out-File -FilePath "$UnblockScriptName" -Append
Write-Output "#>" | Out-File -FilePath "$UnblockScriptName" -Append
Write-Output "Invoke-Command -ComputerName $FileServer -ScriptBlock { ForEach (`$Share in Get-SmbShare) { Unblock-SmbShareAccess -Name `$Share.Name -AccountName `$args[0] -Force } } -Args $ActingUser" | Out-File -FilePath "$UnblockScriptName" -Append
Write-Output "Enable-ADAccount $ActingUser" | Out-File -FilePath "$UnblockScriptName" -Append
Write-Output "Exit" | Out-File -FilePath "$UnblockScriptName" -Append

#Stop Transcriot Logging
Stop-Transcript

#Kill PowerShell process - comment out the line below if you run the script manually through a PowerShell console or the console session will be killed
Stop-Process $PID -Force -Confirm:$false
