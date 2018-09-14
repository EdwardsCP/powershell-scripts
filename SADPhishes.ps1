# SADPhishes.ps1
# Exchange 2016 Compliance Search & Destroy Phishing Emails
# Colin Edwards / @EdwardsCP
# September 2018
#--------------------------------------------------
# Prerequisites: Script must be run from Exchange Management Shell by a User with the Exchange Discovery Management Role
#
# Usage: Execute the script from within EMS
#
# The user is prompted to search using various combinations of the Subject, Sender Address, Date Range, and Attachment Names.
# The user selects to search either All Exchange Locations, or specify the Email Address associated with a specific MailBox or Group
# The script will create and execute a Compliance Search.
# The user then has the option to view details of the search results, delete the Items found by the Search, or Delete the search and exit.
#
# Microsoft's docs say that a Compliance Search will return a max of 500 source mailboxes, and if there are more than 500 mailboxes that contain content that matches the query, the top 500 with the most search results are included in the results.  This means large environments may need to re-run searches.  Look for a future version of this script to be able to loop back through and perform another search if 500 results are returned and then deleted.
#
#=================
#Version 1.0.3
# Added AttachmentNameOptions and AttachmentNameMenu functions to search for emails with a specific Attachment name. 
# Added an option for Attachment Name to the workflow of all searches
# Added an Attachment Name Only search option.
# Added an option for a Pre-Built Suspicious Attachment Types Search, and new functions in that workflow that don't allow for delete. This is for info-gathering only.
#=================
# Version 1.0.2
# Modified ComplianceSearch function to add a TimeStamp to SearchName to make it unique if an identical search is re-run.
# Modified ThisSearchMailboxCount function to display a warning if the Compliance search returns 500 source mailboxes.
#=================
# Version 1.0.1
# Added ThisSearchMailboxCount function to display the number of mailboxes and a list of email addresses with Compliance Search Hits
# Added ExchangeSearchLocationOptions and ExchangeSearchLocationMenu functions so the user can choose to search all Exchange Locations, or limit the search targets based on the Email Address associated with a Mailbox, Distribution Group, or Mail-Enabled Security Group
#=================



#Function to show the full action menu of options
Function MenuOptions{
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the search results."
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
			$PurgeSuffix = "_purge"
			$PurgeName = $SearchName + $PurgeSuffix
			Write-Host "==========================================================================="
			Write-Host "Creating a new Compliance Search Purge Action with the name..."
			Write-Host $PurgeName -ForegroundColor Yellow
			Write-Host "==========================================================================="
			New-ComplianceSearchAction -SearchName "$SearchName" -Purge -PurgeType SoftDelete
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

#Function to show the No Delete action menu of options (for Suspicious Attachment Types Search)
Function NoDeleteMenuOptions{
	Write-host "===================================================="  -ForegroundColor Red
	Write-Host "Take the Search Results above and Investigate." -ForegroundColor Red
	Write-host "===================================================="  -ForegroundColor Red
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the search results."
	Write-Host "[2] Delete this search and Exit."
	}
	
#Function for No Delete menu (for Suspicious Attachment Types Search)
Function ShowNoDeleteMenu{
	Do{
		NoDeleteMenuOptions
		$NoDeleteMenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1 or 2), and press Enter'
		switch ($NoDeleteMenuChoice){
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
			ShowNoDeleteMenu
			}
			
			'2'{
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

# Who doesn't like gratuitous ascii art?
Function DisplayBanner {
	Write-Host ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:                      (:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:( (:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(  :(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:  (:(:(:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:( (:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(: :(:(:  (:(:(:  (:(  :(:(:(  :(: :(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:( (:(:(:(  :(:  (:(:(:(  :(:  (:(:( (:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(: :(:(:(:(:  (  :(:(:(:(:  (  :(:(:(:  (:(:(<><<><<><<><<><  <<><<><<><<><<><<><<><<><<><  <"
	Write-Host ":(:(: :(:(:(:(:(   (:(:(:(:(:(   (:(:(:(:( (:(:(<><<><<><<>    <>    <><<><<><<><<><<><<><<>   <"
	Write-Host ":(:(: :(:(:(:(  :(   (:(:(:(  :(   (:(:(:( (:(:(<><<><<     <><<><<>     ><<><<><><<><<     <<><"
	Write-Host ":(:(: :(:(:(:  (:(:(:  (:(:  (:(:(: :(:(:( (:(:(<><     <<><<><<><<><<><     <<><><     ><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(<    <<><<><<><<><<><<><<><    ><    ><<><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(<><     <<><<><<><<><<><     <<><><     ><<><<><"
	Write-Host ":(:(: :(:(:(:(:(:(:           :(:(:(:(:(:( (:(:(<><<><<     <><<><<>     ><<><<><><<><<     <<><"
	Write-Host ":(:(: :(:(:(:(:(:   :(:(:(:(:   :(:(:(:(:( (:(:(<><<><<><<>    <>    <><<><<><<><<><<><<><<>   <"
	Write-Host ":(:(:( (:(:(:(:   :(:(:(:(:(:(:   :(:(:(: :(:(:(<><<><<><<><<><  <<><<><<><<><<><<><<><<><<><  <"
	Write-Host ":(:(:(: :(:(:(   (:(:(:(:(:(:(:(   (:(:( (:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:( (:(:(: :(:(:(:(:(:(:(:(: :(:(: :(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(: :(:(:(:(:(:(:(:(:(:(:(:(:(:( (:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:( (:(:(:(:(:(:(:(:(:(:(:(:(: :(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:  (:(:(:(:(:(:(:(:(:(:(:  (:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:                       :(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><<><"
	Write-Host "================================================================================================"
	Write-Host "             _____             _____                  _____  _     _     _                 "
	Write-Host "            / ____|    /\     |  __ \                |  __ \| |   (_)   | |                "
	Write-Host "           | (___     /  \    | |  | |               | |__) | |__  _ ___| |__   ___  ___   "
	Write-Host "            \___ \   / /\ \   | |  | |               |  ___/| '_ \| / __| '_ \ / _ \/ __|  "
	Write-Host "            ____) | / ____ \ _| |__| |               | |    | | | | \__ \ | | |  __/\__ \  "
	Write-Host "   _____   |_____(_)_/    \_(_)_____(_)             _|_|_   |_| |_|_|___/_| |_|\___||___/  "
	Write-Host "  / ____|                   | |          ___       |  __ \          | |                    "
	Write-Host " | (___   ___  __ _ _ __ ___| |__       ( _ )      | |  | | ___  ___| |_ _ __ ___  _   _   "
	Write-Host "  \___ \ / _ \/ _' | '__/ __| '_ \      / _ \/\    | |  | |/ _ \/ __| __| '__/ _ \| | | |  "
	Write-Host "  ____) |  __/ (_| | | | (__| | | |    | (_>  <    | |__| |  __/\__ \ |_| | | (_) | |_| |  "
	Write-Host " |_____/ \___|\__,_|_|  \___|_| |_|     \___/\/    |_____/ \___||___/\__|_|  \___/ \__, |  "
	Write-Host "                                                              ________________________/ |  "
	Write-Host "                                                             |@EdwardsCP v1.0.3 2018___/   "
	Write-Host "================================================================================================"
	Write-Host "===============================================================" -ForegroundColor Yellow
	Write-Host "== Exchange 2016 Compliance (S)earch (A)nd (D)estroy Phishes ==" -ForegroundColor Yellow
	Write-Host "===============================================================" -ForegroundColor Yellow
	Write-Host "                                                               " -ForegroundColor Yellow
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "------------------------!!!Warning!!!--------------------------" -ForegroundColor Red
	Write-Host "If you use this script to delete emails, there is no automatic " -ForegroundColor Red
	Write-Host "method to undo the removal of those emails.                    " -ForegroundColor Red
	Write-Host "USE AT YOUR OWN RISK!                                          " -ForegroundColor Red
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "---------------------------------------------------------------" -ForegroundColor Red
	Write-Host "                                                               " -ForegroundColor Yellow
	SearchTypeMenu
	}
	
#Function for SearchType Menu Options Display
Function SearchTypeOptions {
	Write-Host "What type of search are you going to perform?" -ForegroundColor Yellow
	Write-Host "[1] Subject and Sender Address and Date Range"
	Write-Host "[2] Subject and Date Range"
	Write-Host "[3] Subject and Sender Address"
	Write-Host "[4] Subject Only"
	Write-Host "[5] Sender Address Only (DANGEROUS)"
	Write-Host "[6] Attachment Name Only"
	Write-Host "[7] Pre-Built Suspicious Attachment Types Search"
	Write-Host "[Q] Quit"
}

#Function for AttachmentName Menu Options Display
Function AttachmentNameOptions {
	Write-Host "Do you want to search for EMails containing an Attachment with a specific File Name?" -ForegroundColor Yellow
	Write-Host "[1] No"
	Write-Host "[2] Yes"
	Write-Host "[Q] Quit"
}

#Function for AttachmentName Menu
Function AttachmentNameMenu {
	Do{
		AttachmentNameOptions
		$AttachmentNameSelection = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and Press Enter'
		switch ($AttachmentNameSelection){
			'1'{
				ExchangeSearchLocationMenu
			}
			'2'{
				$AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. SADPhishes.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'q'{
			Exit
			}
		}
	}
	until ($AttachmentNameSelection -eq 'q')
}

#Function for ExchangeSearchLocation Menu Options Display
Function ExchangeSearchLocationOptions {
	Write-Host ""
	Write-Host "Do you want to search All Mailboxes, or restrict your search to a specific Mailbox, Distribution Group, or Mail-Enabled Security Group?" -ForegroundColor Yellow
	Write-Host "If you restrict your search, you might leave phishes in other places." -ForegroundColor Yellow
	Write-Host "[1] All Mailboxes"
	Write-Host "[2] A specific MailBox, Distribution Group, or Mail-Enabled Security Group"
	Write-Host "[Q] Quit"
}

#Function for ExchangeSearchLocation Menu
Function ExchangeSearchLocationMenu {
	Do {
		ExchangeSearchLocationOptions
		$ExchangeSearchLocation = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or Q) and press Enter'
		switch ($ExchangeSearchLocation){
			'1'{
				$ExchangeLocation = "all"
				ComplianceSearch
			}
			'2'{
				$ExchangeLocation = Read-Host -Prompt 'Please enter the EMail Address of the MailBox or Group you would like to search within'
				ComplianceSearch
			}
			'q'{
				Exit
			}
		}
	}
	until ($SearchType -eq 'q')
}

#Function for Search Type Menu
Function SearchTypeMenu{
	Do {	
		SearchTypeOptions
		$SearchType = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, 3, 4, 5, 6, 7 or Q) and press Enter'
		switch ($SearchType){
			'1'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$DateRangeSeparator = ".."
				$DateRange = $DateStart + $DateRangeSeparator + $DateEnd
				$ContentMatchQuery = "(Received:$DateRange) AND (From:$Sender) AND (Subject:'$Subject')"
				AttachmentNameMenu
			}
			'2'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$DateStart = Read-Host -Prompt 'Please enter the Beginning Date for your Date Range in the form M/D/YYYY'
				$DateEnd = Read-Host -Prompt 'Please enter the Ending Date for your Date Range in the form M/D/YYYY'
				$DateRangeSeparator = ".."
				$DateRange = $DateStart + $DateRangeSeparator + $DateEnd
				$ContentMatchQuery = "(Received:$DateRange) AND (Subject:'$Subject')"
				AttachmentNameMenu
			}
			'3'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
				$ContentMatchQuery = "(From:$Sender) AND (Subject:'$Subject')"
				AttachmentNameMenu
			}
			'4'{
				$Subject = Read-Host -Prompt 'Please enter the exact Subject of the Email you would like to search for'
				$ContentMatchQuery = "Subject:'$Subject'"
				AttachmentNameMenu
			}
			'5'{
				Do {
					Write-Host "WARNING: Are you sure you want to search based on only Sender Address?" -ForegroundColor Red
					Write-Host "WARNING: This has the potential to return many results and delete many emails." -ForegroundColor Red
					$DangerousSearch = Read-Host -Prompt 'After reading the warning above, would you like to proceed? [Y]es or [Q]uit'
					switch ($DangerousSearch){
						'Y'{
							$Sender = Read-Host -Prompt 'Please enter the exact Sender (From:) address of the Email you would like to search for'
							$ContentMatchQuery = "From:$Sender"
							AttachmentNameMenu
						}
						'q'{
							Exit
						}
					}
				}
				until ($DangerousSearch -eq 'q')
			}
			'6'{
				$AttachmentName = Read-Host -Prompt 'Please enter the exact File Name of the Attachment you want to search for (i.e. SADPhishes.ps1) and Press Enter'
				ExchangeSearchLocationMenu
			}
			'7'{
				Write-Host "You have chosen to conduct the SADPhishes Pre-Built Suspicious Attachment Types Search." -ForegroundColor Yellow
				Write-Host "This search will return a list of Mailboxes that contain Attachments with specific file extensions." -ForegroundColor Yellow
				Write-Host "This search is a Search-Only option, with no Delete built into the SADPhishes Workflow." -ForegroundColor Yellow
				Write-Host "Take these results and investigate." -ForegroundColor Yellow
				Read-Host -Prompt "After you have read the information about this Suspicious Attachment Search, Press Enter to continue."
				$ContentMatchQuery = "((HasAttachment:true) AND (Attachment:'.ade') OR (Attachment:'.adp') OR (Attachment:'.apk') OR (Attachment:'.bas') OR (Attachment:'.bat') OR (Attachment:'.chm') OR  (Attachment:'.cmd') OR (Attachment:'.com') OR (Attachment:'.cpl') OR (Attachment:'.dll') OR (Attachment:'.exe') OR (Attachment:'.hta') OR (Attachment:'.inf') OR (Attachment:'.iqy') OR (Attachment:'.jar') OR (Attachment:'.js') OR (Attachment:'.jse') OR (Attachment:'.lnk') OR (Attachment:'.msc') OR (Attachment:'.msi') OR (Attachment:'.msp') OR (Attachment:'.mst') OR (Attachment:'.ocx') OR (Attachment:'.pif') OR (Attachment:'.pl') OR (Attachment:'.ps1') OR (Attachment:'.reg') OR (Attachment:'.scr') OR (Attachment:'.sct') OR (Attachment:'.shs') OR (Attachment:'.slk') OR (Attachment:'.sys') OR (Attachment:'.vb') OR (Attachment:'.vbe') OR (Attachment:'.vbs') OR (Attachment:'.wsc') OR (Attachment:'.wsf') OR (Attachment:'.wsh'))"
				ExchangeSearchLocationMenu
			}
			'q'{
				Exit
			}
		}
	}
	until ($SearchType -eq 'q')
}

#Function to count and list Mailboxes with Search Hits.  Code mostly taken from a MS TechNet article.
Function ThisSearchMailboxCount {
	$ThisSearchResults = $ThisSearch.SuccessResults;
	if (($ThisSearch.Items -le 0) -or ([string]::IsNullOrWhiteSpace($ThisSearchResults))){
               Write-Host "!!!The Compliance Search didn't return any useful results!!!" -ForegroundColor Red
	}
	$mailboxes = @() #create an empty array for mailboxes
	$ThisSearchResultsLines = $ThisSearchResults -split '[\r\n]+'; #Split up the Search Results at carriage return and line feed
	foreach ($ThisSearchResultsLine in $ThisSearchResultsLines){
		# If the Search Results Line matches the regex, and $matches[2] (the value of "Item count: n") is greater than 0)
		if ($ThisSearchResultsLine -match 'Location: (\S+),.+Item count: (\d+)' -and $matches[2] -gt 0){ 
			# Add the Location: (email address) for that Search Results Line to the $mailboxes array
			$mailboxes += $matches[1]; 
		}
	}
	Write-Host "Number of mailboxes that have Search Hits..."
	Write-Host $mailboxes.Count -ForegroundColor Yellow
	Write-Host "List of mailboxes that have Search Hits..."
	write-Host $mailboxes -ForegroundColor Yellow
	if ($mailboxes.Count -gt 499) {
		Write-Host "============WARNING - There are 500 or more Mailboxes with results!============" -ForegroundColor Red
		Write-Host "Microsoft's Compliance Search can search everywhere, but only returns the top" -ForegroundColor Red
		Write-Host "500 Mailboxes with the most hits that match the search!" -ForegroundColor Red
		Write-Host " " 
		Write-Host "If you use this search to delete Email Items, you will need to run the same" -ForegroundColor Red
		Write-Host "query again to return more mailboxes if there are more than 500 with hits." -ForegroundColor Red
		Read-Host -Prompt "Please press Enter after reading the warning above."
	}
}

Function ComplianceSearch {
	#Set SearchName based on SearchType
	switch ($SearchType){
			'1'{
				$SearchName = "Remove Subject [$Subject] Sender [$Sender] DateRange [$DateRange] ExchangeLocation [$ExchangeLocation] Phishing Message"
			}
			'2'{
				$SearchName = "Remove Subject [$Subject] DateRange [$DateRange] ExchangeLocation [$ExchangeLocation] Phishing Message"
			}
			'3'{
				$SearchName = "Remove Subject [$Subject] Sender [$Sender] ExchangeLocation [$ExchangeLocation] Phishing Message"
			}
			'4'{
				$SearchName = "Remove Subject [$Subject] ExchangeLocation [$ExchangeLocation] Phishing Message"
			}
			'5'{
				$SearchName = "Remove Sender [$Sender] ExchangeLocation [$ExchangeLocation] Phishing Message"
			}
			'6'{
				$SearchName = "Remove Exchange Location [$ExchangeLocation] Phishing Message"
			}
			'7'{
				$SearchName = "SADPhishes Pre-Built Suspicious Attachment Types Search Exchange Location [$ExchangeLocation]"
			}
	}
	#If an AttachmentName has been specified, append it to $SearchName and modify $ContentMatchQuery to specify the email has an attachment and what the name is.
	if ($AttachmentName -ne $null){
		$SearchName = $SearchName + " with Attachment [" + $AttachmentName + "]"
		If ($ContentMatchQuery -ne $null){
		$ContentMatchQuery = "(HasAttachment:true) AND (Attachment:'$AttachmentName') AND " + $ContentMatchQuery
		}
		If ($ContentMatchQuery -eq $null){
		$ContentMatchQuery = "(HasAttachment:true) AND (Attachment:'$AttachmentName')"
		}
	}	
	
	# Timestamp the SearchName (to make it unique), then Create and Execute a New Compliance Search based on the user set Variables
	$TimeStamp = Get-Date -Format o | foreach {$_ -replace ":", "."}
	$SearchName = $SearchName + " " + $TimeStamp
	Write-Host "==========================================================================="
	Write-Host "Creating a new Compliance Search with the name..."
	Write-Host $SearchName -ForegroundColor Yellow
	Write-Host "...using the query..."
	Write-Host $ContentMatchQuery -ForegroundColor Yellow
	Write-Host "==========================================================================="
	
	New-ComplianceSearch -Name "$SearchName" -ExchangeLocation $ExchangeLocation -ContentMatchQuery $ContentMatchQuery
	Start-ComplianceSearch -Identity "$SearchName"
	Get-ComplianceSearch -Identity "$SearchName"
	#Display status, then results of Compliance Search
	do{
		$ThisSearch = Get-ComplianceSearch -Identity $SearchName
		Start-Sleep 2
		Write-Host $ThisSearch.Status
	}
	until ($ThisSearch.status -match "Completed")

	Write-Host "==========================================================================="
	Write-Host The search returned...
	Write-Host $ThisSearch.Items Items -ForegroundColor Yellow
	Write-Host That match the query...
	Write-Host $ContentMatchQuery -ForegroundColor Yellow
	ThisSearchMailboxCount
	Write-Host "==========================================================================="
	#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
	if ($SearchType -match "7"){
		ShowNoDeleteMenu
	}
	
	ShowMenu
}

#Drop the user into the DisplayBanner function (and then Search Type Menu) to begin the process.
DisplayBanner