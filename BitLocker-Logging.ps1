# BitLocker-Logging.ps1
# Colin Edwards
# August 2018
#-------------------------
# The purpose of this script is to log the BitLocker status of a computer.
# A log file is created for each computer hostname, with a new file created for every new month, in the format "YYYY-MM-HOSTNAME.csv"
# Configure this as a Group Policy Shutdown/Startup script or use Task Scheduler or some other method to get the computers on your domain to run this regularly.

$blvs = Get-BitLockerVolume
$CurrentDateTime = Get-Date -Format o
$CurrentYearMonth = Get-Date -Format yyyy-MM
$LogFileName = "$CurrentYearMonth-$env:COMPUTERNAME.csv"
$LogFilePath = "\\Server\BitLockerStatus$"
$LogFile = "$LogFilePath\$LogFileName"

# If the log file for the current Year and Month does not exist, then create the log file and insert a header row
if (!(Test-Path $LogFile)) {
	Set-Content $LogFile -Value "Timestamp,Hostname,Drive,BitLockerStatus"
}

# If the log file for the current Year and Month does exist, then append a new line to it with details for each BitLocker Volume (drive) in the workstation 
if (Test-Path $LogFile) {
	foreach ($blv in $blvs) {
	$blvvolumestatus = $blv.volumestatus
	$blvMountPoint = $blv.MountPoint
	$CurrentLog = "$CurrentDateTime,$env:COMPUTERNAME,$blvMountPoint,$blvvolumestatus"
	Out-File -filepath $LogFile -append -inputobject $CurrentLog -encoding ASCII
	}
}

exit
