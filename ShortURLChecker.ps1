# ShortURLChecker.ps1
# Powershell to find where a shortened URL redirects
# Colin Edwards / @EdwardsCP

Write-Host "Running ShortURLChecker.ps1" -ForegroundColor Yellow
Write-Host "===========================" -ForegroundColor Yellow
$script:shorturl = Read-Host -Prompt 'Please enter the Shortened URL you want to check.'
Do {
	$script:CheckURL = Invoke-WebRequest -uri $script:shorturl -MaximumRedirection 0 -EA SilentlyContinue
	$script:CheckURLRaw = $script:CheckURL.RawContent
	$script:CheckURLLocation = $script:CheckURLRaw -match 'Location: (.*)'
	$script:LongURL = $Matches[1]
	Write-Host "===========================" -ForegroundColor Yellow
	Write-Host "The URL..." -ForegroundColor Yellow
	Write-Host $Script:shorturl
	Write-Host "...was redirected to this location..." -ForegroundColor Yellow
	Write-Host $Script:LongURL
	Write-Host "...checking that URL for another redirect..." -ForegroundColor Yellow
	Write-Host "===========================" -ForegroundColor Yellow
	$script:LongURLCheck = Invoke-WebRequest -uri $script:LongURL -MaximumRedirection 0 -EA SilentlyContinue
	$Script:LongURLCheckRaw = $Script:LongURLcheck.RawContent
	$Script:LongURLCheckLocation = $script:LongURLCheckRaw -match 'Location: (.*)'
	$script:LongURLLocation = $Matches[1]
	If ($Script:LongURLLocation -eq $script:LongURL){
		Write-Host "===========================" -ForegroundColor Yellow
		Write-Host "No more redirects found." -ForegroundColor Yellow
		Write-Host "This appears to be the final URL..." -ForegroundColor Yellow
		Write-Host $script:LongURL
		Write-Host "No more redirects found" -ForegroundColor Green
		Write-Host "No more redirects found" -ForegroundColor Green
		Write-Host "No more redirects found" -ForegroundColor Green
		Write-Host "===========================" -ForegroundColor Yellow
		exit
		}
	If ($Script:LongURLLocation -ne $Script:LongURL){
		Write-Host "===========================" -ForegroundColor Yellow
		Write-Host "Finding the next redirect..." -ForegroundColor Yellow
		$Script:ShortURL = $Script:LongURL
		}
}
Until ($Script:LongURLLocation -eq $script:LongURL)
