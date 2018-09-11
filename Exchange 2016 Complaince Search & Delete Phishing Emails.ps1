# Exchange 2016 Compliance Search & Delete Phishing Emails
# Colin Edwards / @EdwardsCP
# September 2018
#--------------------------------------------------
# Usage:
# User is prompted for the exact Subject and Sender.
# The script will create and execute a Compliance Search.
# User then has the option to view details of the search results,
# delete the Items found by the Search, or Delete the search
# and exit.


#Function to show the full action menu of options
Function MenuOptions{
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the search results so you can review the location of these Items."
	Write-Host "[2] Delete the Items (move them to Deleted Recoverable Items)."
	Write-Host "[3] Delete this search and Exit."
	}

#Function for full action menu
Function ShowMenu{
	Do{
		MenuOptions
		$MenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or 3), and press Enter'
		switch ($MenuChoice){
			'1'{
			$ThisSearch | Format-List
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-Host "Please review the output above" -ForegroundColor Red
			Write-host "After reviewing, please make another selection below"  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			ShowMenu
			}
			
			'2'{
			New-ComplianceSearchAction -SearchName "$SearchName" -Purge -PurgeType SoftDelete
			$PurgeSuffix = "_purge"
			$PurgeName = $SearchName + $PurgeSuffix
				do{
				$ThisPurge = Get-ComplianceSearchAction -Identity $PurgeName
				Start-Sleep 2
				Write-Host $ThisPurge.Status
				}
				until ($ThisPurge.Status -match "Completed")
			$ThisPurge | Format-List
			Write-Host "The items have been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to exit"
			Exit
			}
		
			'3'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to exit"
			Exit
			}
			'q'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to exit"
			Exit
			}
		}
	}
	Until ($MenuChoice -eq 'q')
}


#Function for SearchType Menu Options
Function SearchTypeOptions {
	Write-Host "==============================================================" -ForegroundColor Yellow
	Write-Host "== Exchange 2016 Compliance Search & Delete Phishing Emails ==" -ForegroundColor Yellow
	Write-Host "==============================================================" -ForegroundColor Yellow
	Write-Host "=========================================|@EdwardsCP|09-2018|=" -ForegroundColor Yellow
	Write-Host "==============================================================" -ForegroundColor Yellow
	Write-Host "                                                              " -ForegroundColor Yellow
	Write-Host "--------------------------------------------------------------" -ForegroundColor Red
	Write-Host "------------------------!!!Warning!!!-------------------------" -ForegroundColor Red
	Write-Host "If you use this script to delete emails, there is no automatic" -ForegroundColor Red
	Write-Host "method to undo the removal of those emails.                   " -ForegroundColor Red
	Write-Host "USE ARE YOUR OWN RISK!                                        " -ForegroundColor Red
	Write-Host "--------------------------------------------------------------" -ForegroundColor Red
	Write-Host "--------------------------------------------------------------" -ForegroundColor Red
	Write-Host "                                                              " -ForegroundColor Yellow
	Write-Host "What type of search are you going to perform?" -ForegroundColor Yellow
	Write-Host "[1] Subject and Sender Address and Date Range"
	Write-Host "[2] Subject and Date Range"
	Write-Host "[3] Subject and Sender Address"
	Write-Host "[4] Subject Only"
	Write-Host "[Q] Quit"
}

#Function for Search Type Menu
Function SearchTypeMenu{
	Do {	
		SearchTypeOptions
		$SearchType = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and press Enter'
		switch ($SearchType){
			'1'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$DateRangeSeparator = ".."
				$DateRange = $DateStart + $DateRangeSeparator + $DateEnd
				$ContentMatchQuery = "(Received:$DateRange) AND (From:$Sender) AND (Subject:'$Subject')"
				ComplianceSearch
			}
			'2'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$DateRangeSeparator = ".."
				$DateRange = $DateStart + $DateRangeSeparator + $DateEnd
				$ContentMatchQuery = "(Received:$DateRange) AND (Subject:'$Subject')"
				ComplianceSearch
			}
			'3'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$ContentMatchQuery = "(From:$Sender) AND (Subject:'$Subject')"
				ComplianceSearch
			}
			'4'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$ContentMatchQuery = "Subject:'$Subject'"
				ComplianceSearch
			}

			'q'{
				Exit
			}
		}
	}
	until ($SearchType -eq 'q')
}


Function ComplianceSearch {
	#Set Variables, prompt user for input
	$SearchName = "Remove $Subject Phishing Message"
	
	#Create and Execute a New Compliance Search based on the user set Variables
	New-ComplianceSearch -Name "$SearchName" -ExchangeLocation all -ContentMatchQuery $ContentMatchQuery
	Start-ComplianceSearch -Identity "$SearchName"
	
	#Display status, then results of Compliance Search
	do{
		$ThisSearch = Get-ComplianceSearch -Identity $SearchName
		Start-Sleep 2
		Write-Host $ThisSearch.Status
	}
	until ($ThisSearch.status -match "Completed")

	Write-Host ===========================================================================
	Write-Host The search returned...
	Write-Host $ThisSearch.Items Items -ForegroundColor Yellow
	Write-Host That match the query...
	Write-Host $ContentMatchQuery -ForegroundColor Yellow
	Write-Host ===========================================================================

	ShowMenu
}

#Drop the user into the Search Type Menu to begin the process.
SearchTypeMenu

