# SADPhishes.ps1
# Exchange 2016 Compliance Search & Destroy Phishing Emails
# Colin Edwards / @EdwardsCP
# September 2018
#--------------------------------------------------
# Prerequisites: Script must be run from Exchange Management Shell by a User with the Exchange Discovery Management Role
#
# Basic Usage: 	Execute the script from within EMS
#			The user is prompted to search using various combinations of the Subject, Sender Address, Date Range, and Attachment Names.
#			The user selects to search either All Exchange Locations, or specify the Email Address associated with a specific MailBox or Group
#			The script will create and execute a Compliance Search.
#			The user then has the option to view details of the search results, delete the Items found by the Search, create an eDiscovery Search or Delete the search and return to the mail search options menu.
#
#
# Microsoft's docs say that a Compliance Search will return a max of 500 source mailboxes, and if there are more than 500 mailboxes that contain content that matches the query, the top 500 with the most search results are included in the results.  This means large environments may need to re-run searches.  Look for a future version of this script to be able to loop back through and perform another search if 500 results are returned and then deleted.
#
#=================
#Version 1.0.4
# Added options to create an Exchange In-place eDiscovery Search from the Compliance Search results.
# The option to execute the eDiscovery search is completely experimental. It knocked Exchange offline during testing. Not recommended in Prod.
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
	Write-host "===================================================="
	Write-Host "COMPLIANCE SEARCH ACTIONS MENU" -ForegroundColor Green
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the Compliance Search results."
	Write-Host "[2] Delete the Items (move them to Deleted Recoverable Items). WARNING: No automated way to restore them!"
	Write-Host "[3] Create an Exchange In-Place eDiscovery Search from the Compliance Search results."
	Write-Host "[4] Delete this search and Return to the Search Options Menu."
	}
	
#Function to show the eDiscovery Search Action menu of options
Function EDiscoverySearchMenuOptions{
	Write-host "===================================================="
	Write-Host "EDISCOVERY SEARCH ACTIONS MENU" -ForegroundColor Green
	Write-host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the new In-Place eDiscovery Search."
	Write-Host "[2] Start the new In-Place eDiscovery Search. (WARNING: EXPERIMENTAL! Not recommended for Production.)"
	Write-Host "[3] Delete the new In-Place eDiscovery Search and return to the Compliance Search Actions Menu."
	}

#Function for the eDiscovery Search Action menu
Function ShowEDiscoverySearchMenu {
	EDiscoverySearchMenuOptions
	$EDiscoverySearchMenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2, or 3), and press Enter'
	Switch ($EDiscoverySearchMenuChoice){
		'1'{
			$ThisEDiscoverySearch | Format-List
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-Host "Please review the output above" -ForegroundColor Red
			Write-host "After reviewing, please make another selection below"  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			Write-host "===================================================="  -ForegroundColor Red
			ShowEDiscoverySearchMenu
		}
		
		'2'{
			Do {
				Write-Host "WARNING: Executing an In-Place eDiscovery Search created by SADPhishes is EXPERIMENTAL!" -ForegroundColor Red
				Write-Host "WARNING: This has been known to generate errors and knock the Mailbox Server offline!" -ForegroundColor Red
				Write-Host "WARNING: This action is NOT recommended for Production use." -ForegroundColor Red
				Write-Host "You have been warned." -ForegroundColor Red
				$DangerousEDiscoverySearch = Read-Host -Prompt 'After reading the warning above, would you like to proceed with executing the search? [Y]es or [Q]uit'
				switch ($DangerousEDiscoverySearch){
					'Y'{
						Write-Host "This might blow up.  You're on your own to clean up the mess." -ForegroundColor Red
						Write-Host "==========================================================================="
						Write-Host "Starting the new In-Place eDiscovery Search with the name..."
						Write-Host "$ThisEDiscoverySearchName" -ForegroundColor Yellow
						Write-Host "...that will search against these mailboxes..."
						Write-Host $mailboxes -ForegroundColor Yellow
						Write-Host "...using the Search Query..."
						Write-Host $ContentMatchQuery -ForegroundColor Yellow
						Write-Host "==========================================================================="
						Write-Host "Please wait for the In-Place eDiscovery Search to complete..." -ForegroundColor Yellow
						Write-Host "==========================================================================="
						Start-MailboxSearch -Identity $ThisEDiscoverySearchName
							do{
							$ThisEDiscoverySearch = Get-MailboxSearch $ThisEDiscoverySearchName
							Start-Sleep 2
							Write-Host $ThisEDiscoverySearch.Status
							}
							until ($ThisEDiscoverySearch.Status -match "EstimateSucceeded")
						Write-Host "==========================================================================="
						Write-Host "The In-Place eDiscovery Search has completed."
						Write-Host "You can use this URL to Preview the Results..." 
						Write-Host $ThisEDiscoverySearch.PreviewResultsLink -ForegroundColor Yellow
						Write-Host "If you need to Copy those results to a Discovery Mailbox, or Export them"
						Write-Host "to a PST file, please use Exchange Administrative Center's Compliance "
						Write-Host "Management In-Place eDiscovery workflow to proceed with those actions."
						Write-Host "==========================================================================="
						Write-Host " "
						Read-Host -Prompt "Please review all of the information above and then press Enter to proceed."
						Write-Host "SADPhishes will now return to the Compliance Search Actions menu where you" -ForegroundColor Yellow
						Write-Host "will have the option to delete all of the emails with Search Hits." -ForegroundColor Yellow
						Write-Host "The eDiscovery Searches that were created during this session are not being" -ForegroundColor Yellow
						Write-Host "deleted." -ForegroundColor Yellow
						Read-Host "Please review all of the information above then press Enter to return to the Compliance Search Actions Menu."
						#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
						if ($SearchType -match "7"){
							ShowNoDeleteMenu
						}
						#If the search was any other type, show the regular Actions menu that allows Delete.
						ShowMenu
					}
					'q'{
						Write-Host "Proceeding to return to the Compliance Search Actions Menu..."
						Do{
							Write-Host "==========================================================================="
							Write-Host "Do you want to Remove the new In-Place eDiscovery Search with the name..."
							Write-Host "$ThisEDiscoverySearchName" -ForegroundColor Yellow
							Write-Host "...or do you want to leave it in place?"
							Write-Host "[1] Delete the eDiscovery Search and return to the Compliance Search Actions Menu."
							Write-Host "[2] Return to the Compliance Search Actions Menu without deleting."
							$DangerousEDiscoverySearchQuitChoice = Read-Host -Prompt 'Please enter a selection from the menu (1 or 2) and press Enter.'
							switch ($DangerousEDiscoverySearchQuitChoice){
								'1'{
									Remove-MailboxSearch -Identity $ThisEDiscoverySearchName
									Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
									Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
									#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
									if ($SearchType -match "7"){
										ShowNoDeleteMenu
									}
									#If the search was any other type, show the regular Actions menu that allows Delete.
									ShowMenu
								}
								'2'{
									#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
									if ($SearchType -match "7"){
										ShowNoDeleteMenu
									}
									#If the search was any other type, show the regular Actions menu that allows Delete.
									ShowMenu
								}
							}
						}
						Until ($DangerousEDiscoverySearchQuitChoice -eq '1')
					}
				}
			}
		
		until ($DangerousEDiscoverySearch -eq 'q')
		}
		
		'3'{
			Remove-MailboxSearch -Identity $ThisEDiscoverySearchName
			Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
			#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
			if ($SearchType -match "7"){
				ShowNoDeleteMenu
			}
			#If the search was any other type, show the regular Actions menu that allows Delete.
			ShowMenu
		}
		
		'q'{
			Remove-MailboxSearch -Identity $ThisEDiscoverySearchName
			Write-Host "The eDiscovery Search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to return to the Compliance Search Actions Menu"
			#If the search was a Pre-Built Suspicious Attachment Types Search, don't give the user the regular Actions menu that allows them to Delete.
			if ($SearchType -match "7"){
				ShowNoDeleteMenu
			}
			#If the search was any other type, show the regular Actions menu that allows Delete.
			ShowMenu	
		}
	}
	Until ($EDiscoverySearchMenuChoice -eq 'q')
}
	
#Function to create an eDiscovery Search. Code mostly taken from a MS TechNet article.  
 Function CreateEDiscoverySearch{
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
	#Name the the EDiscoverySearch (MailboxSearch) using the Compliance Search's name, followed by _MBSearch, followed by an integer. increase the integer until you hit a name that doesn't already exist.
	$EDiscoverySearchName = $SearchName + "_MBSearch";
	$I = 1;
	$MailboxSearches = Get-MailboxSearch;
		while ($true){
			$found = $false
			$ThisEDiscoverySearchRun = "$EDiscoverySearchName$I"
			foreach ($MailboxSearch in $MailboxSearches){
				if ($MailboxSearch.Name -eq $ThisEDiscoverySearchRun){
					$found = $true;
					break;
				}
		}
		if (!$found){
			break;
		}
		$I++;
		}
	$ThisEDiscoverySearchName = "$EDiscoverySearchName$i"
	Write-Host "==========================================================================="
	Write-Host "Creating a new In-Place eDiscovery Search with the name..."
	Write-Host "$ThisEDiscoverySearchName" -ForegroundColor Yellow
	Write-Host "...that will search against these mailboxes..."
	Write-Host $mailboxes -ForegroundColor Yellow
	Write-Host "...using the Search Query..."
	Write-Host $ContentMatchQuery -ForegroundColor Yellow
	Write-Host "==========================================================================="
	New-MailboxSearch "$ThisEDiscoverySearchName" -SourceMailboxes $mailboxes -SearchQuery $ContentMatchQuery -EstimateOnly
	$ThisEDiscoverySearch = Get-MailboxSearch $ThisEDiscoverySearchName
	do{
		$ThisEDiscoverySearch = Get-MailboxSearch $ThisEDiscoverySearchName
		Start-Sleep 1
	}
	Until ($ThisEDiscoverySearch -ne $null)
	Write-Host "New In-Place eDiscovery Search Successfully Created!" -ForegroundColor Yellow
	ShowEDiscoverySearchMenu
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
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				Write-Host "WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!" -ForegroundColor Red
				$DangerousPurge = Read-Host -Prompt 'After reading the warning above, would you like to proceed? [Y]es or [Q]uit'
				Do {
					switch ($DangerousPurge){
						'Y'{
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
							Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
							ClearSADPhishesVars
							SearchTypeMenu
						}
						'q'{
							Read-Host -Prompt "Please press Enter to return to the Compliance Search Actions Menu"
							ShowMenu
						}
					}
				}
				Until ($DangerousPurge -eq 'q')
			#$PurgeSuffix = "_purge"
			#$PurgeName = $SearchName + $PurgeSuffix
			#Write-Host "==========================================================================="
			#Write-Host "Creating a new Compliance Search Purge Action with the name..."
			#Write-Host $PurgeName -ForegroundColor Yellow
			#Write-Host "==========================================================================="
			#New-ComplianceSearchAction -SearchName "$SearchName" -Purge -PurgeType SoftDelete
			#	do{
			#	$ThisPurge = Get-ComplianceSearchAction -Identity $PurgeName
			#	Start-Sleep 2
			#	Write-Host $ThisPurge.Status
			#	}
			#	until ($ThisPurge.Status -match "Completed")
			#$ThisPurge | Format-List
			#Write-Host "The items have been deleted." -ForegroundColor Red
			#Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			#ClearSADPhishesVars
			#SearchTypeMenu
			}
			
			'3'{
			CreateEDiscoverySearch
			}
			'4'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
			
			'q'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
		}
	}
	Until ($MenuChoice -eq 'q')
}

#Function to show the No Delete action menu of options (for Suspicious Attachment Types Search)
Function NoDeleteMenuOptions{
	Write-Host "COMPLIANCE SEARCH ACTIONS MENU (No Delete)" -ForegroundColor Green
	Write-Host "Note: As a precaution, the delete option is not available for a Suspicious Attachment Types Search." -ForegroundColor Yellow
	Write-Host How would you like to proceed?
	Write-Host "[1] Display the Detailed (Format-List) view of the search results."
	Write-Host "[2] Create an Exchange In-Place eDiscovery Search from the results."
	Write-Host "[3] Delete this search and Return to the Search Options Menu."
	}
	
#Function for No Delete menu (for Suspicious Attachment Types Search)
Function ShowNoDeleteMenu{
	Do{
		NoDeleteMenuOptions
		$NoDeleteMenuChoice = Read-Host -Prompt 'Please enter a selection from the menu (1, 2 or 3) and press Enter'
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
			CreateEDiscoverySearch
			}
			
			'3'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
			}
			
			'q'{
			Remove-ComplianceSearch -Identity $SearchName
			Write-Host "The search has been deleted." -ForegroundColor Red
			Read-Host -Prompt "Press Enter to Return to the Search Options Menu"
			ClearSADPhishesVars
			SearchTypeMenu
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
	Write-Host "                                                             |@EdwardsCP v1.0.5 2018___/   "
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
	ClearSADPhishesVars
	SearchTypeMenu
	}
	
#Function for SearchType Menu Options Display
Function SearchTypeOptions {
	Write-Host "SEARCH OPTIONS MENU" -ForegroundColor Green
	Write-Host "What type of search are you going to perform?" -ForegroundColor Yellow
	Write-Host "[1] Subject and Sender Address and Date Range"
	Write-Host "[2] Subject and Date Range"
	Write-Host "[3] Subject and Sender Address"
	Write-Host "[4] Subject Only"
	Write-Host "[5] Sender Address Only (DANGEROUS)"
	Write-Host "[6] Attachment Name Only"
	Write-Host "[7] Pre-Built Suspicious Attachment Types Search"
	Write-Host "[Q] Quit"
	Write-host "---Debugging Options---"
	Write-Host "[X] gci: variable"
	Write-Host "[Y] Print SADPhishesVars"
	Write-Host "[Z] Clear SADPhishesVars"
}

#Function for AttachmentName Menu Options Display
Function AttachmentNameOptions {
	Write-Host "ATTACHMENT OPTIONS MENU" -ForegroundColor Green
	Write-Host "Do you want to search for EMails containing an Attachment with a specific File Name?" -ForegroundColor Yellow
	Write-Host "[1] No"
	Write-Host "[2] Yes"
	Write-Host "[Q] Quit and Return to the Search Options Menu"
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
				ClearSADPhishesVars
				SearchTypeMenu
			}
		}
	}
	until ($AttachmentNameSelection -eq 'q')
}

#Function for ExchangeSearchLocation Menu Options Display
Function ExchangeSearchLocationOptions {
	Write-Host ""
	Write-Host "LOCATION OPTIONS MENU" -ForegroundColor Green
	Write-Host "Do you want to search All Mailboxes, or restrict your search to a specific Mailbox, Distribution Group, or Mail-Enabled Security Group?" -ForegroundColor Yellow
	Write-Host "If you restrict your search, you might leave phishes in other places." -ForegroundColor Yellow
	Write-Host "[1] All Mailboxes"
	Write-Host "[2] A specific MailBox, Distribution Group, or Mail-Enabled Security Group"
	Write-Host "[Q] Quit and Return to the Search Options Menu"
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
				ClearSADPhishesVars
				SearchTypeMenu
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
				$ContentMatchQuery = "(Subject:'$Subject')"
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
							$ContentMatchQuery = "(From:$Sender)"
							AttachmentNameMenu
						}
						'q'{
							Read-Host -Prompt "Please press Enter to return to the Search Options Menu"
							ClearSADPhishesVars
							SearchTypeMenu
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
				#Note: removed (HasAttachment:true) property for troubleshooting failed eDiscovery Searches further in the workflow.
				#$ContentMatchQuery = "((HasAttachment:true) AND (Attachment:'.ade') OR (Attachment:'.adp') OR (Attachment:'.apk') OR (Attachment:'.bas') OR (Attachment:'.bat') OR (Attachment:'.chm') OR  (Attachment:'.cmd') OR (Attachment:'.com') OR (Attachment:'.cpl') OR (Attachment:'.dll') OR (Attachment:'.exe') OR (Attachment:'.hta') OR (Attachment:'.inf') OR (Attachment:'.iqy') OR (Attachment:'.jar') OR (Attachment:'.js') OR (Attachment:'.jse') OR (Attachment:'.lnk') OR (Attachment:'.msc') OR (Attachment:'.msi') OR (Attachment:'.msp') OR (Attachment:'.mst') OR (Attachment:'.ocx') OR (Attachment:'.pif') OR (Attachment:'.pl') OR (Attachment:'.ps1') OR (Attachment:'.reg') OR (Attachment:'.scr') OR (Attachment:'.sct') OR (Attachment:'.shs') OR (Attachment:'.slk') OR (Attachment:'.sys') OR (Attachment:'.vb') OR (Attachment:'.vbe') OR (Attachment:'.vbs') OR (Attachment:'.wsc') OR (Attachment:'.wsf') OR (Attachment:'.wsh'))"
				$ContentMatchQuery = "((Attachment:'.ade') OR (Attachment:'.adp') OR (Attachment:'.apk') OR (Attachment:'.bas') OR (Attachment:'.bat') OR (Attachment:'.chm') OR (Attachment:'.cmd') OR (Attachment:'.com') OR (Attachment:'.cpl') OR (Attachment:'.dll') OR (Attachment:'.exe') OR (Attachment:'.hta') OR (Attachment:'.inf') OR (Attachment:'.iqy') OR (Attachment:'.jar') OR (Attachment:'.js') OR (Attachment:'.jse') OR (Attachment:'.lnk') OR (Attachment:'.mht') OR (Attachment:'.msc') OR (Attachment:'.msi') OR (Attachment:'.msp') OR (Attachment:'.mst') OR (Attachment:'.ocx') OR (Attachment:'.pif') OR (Attachment:'.pl') OR (Attachment:'.ps1') OR (Attachment:'.reg') OR (Attachment:'.scr') OR (Attachment:'.sct') OR (Attachment:'.shs') OR (Attachment:'.slk') OR (Attachment:'.sys') OR (Attachment:'.vb') OR (Attachment:'.vbe') OR (Attachment:'.vbs') OR (Attachment:'.wsc') OR (Attachment:'.wsf') OR (Attachment:'.wsh'))"
				ExchangeSearchLocationMenu
			}
			'q'{
				Write-Host "Thanks for using SADPhishes!" -ForegroundColor Yellow
				Read-Host -Prompt "Please press Enter to exit."
				Exit
			}
			'x'{
			gci variable:
			}
			'y'{
			PrintSADPhishesVars
			}
			'z'{
			ClearSADPhishesVars
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
	#If an AttachmentName has been specified, Modify SearchName to include it.  
	if ($AttachmentName -ne $null){
		$SearchName = $SearchName + " with Attachment [" + $AttachmentName + "]"
		# If a ContentMatchQuery is already set, modify $ContentMatchQuery to include the attachment.
		If ($ContentMatchQuery -ne $null){
		#Note: removed (HasAttachment:true) property for troubleshooting failed eDiscovery Searches further in the workflow.
		#$ContentMatchQuery = "(HasAttachment:true) AND (Attachment:'$AttachmentName') AND " + $ContentMatchQuery
		$ContentMatchQuery = "(Attachment:'$AttachmentName') AND " + $ContentMatchQuery
		}
		# If an AttachmentName has been specified, and a ContentMatchQuery is NOT already set, set the ContentMatchQuery.
		If ($ContentMatchQuery -eq $null){
		#Note: removed (HasAttachment:true) property for troubleshooting failed eDiscovery Searches further in the workflow.
		#$ContentMatchQuery = "(HasAttachment:true) AND (Attachment:'$AttachmentName')"
		$ContentMatchQuery = "(Attachment:'$AttachmentName')"
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
	#If a Subject was specified, warn the user about Microsoft returning results with additional text before or after the subject that was defined.
	if ($Subject -ne $null){
		Write-Host "===========================================================================" -ForegroundColor Yellow
		Write-Host "Warning: Your Compliance Search contained a Subject [$Subject]."             -ForegroundColor Yellow
		Write-Host "When you use the Subject property in a query, the search returns all"        -ForegroundColor Yellow
		Write-Host "messages in which the subject line contains the text you are searching for." -ForegroundColor Yellow
		Write-Host "The query doesn't only return exact matches.  For example, if you search"    -ForegroundColor Yellow
		Write-Host "(Subject:SADPhishes), your results will include messages with the subject"   -ForegroundColor Yellow
		Write-Host "'SADPhishes', but also messages with the subjects 'SADPhishes is good!' and" -ForegroundColor Yellow
		Write-Host "'RE: Screw SADPhishes. it sucks!'"                                           -ForegroundColor Yellow
		Write-Host " "                                                                           -ForegroundColor Yellow
		Write-Host "This is just how the Microsoft Exchange Content Search works."               -ForegroundColor Yellow
		Write-Host " "                                                                           -ForegroundColor Yellow
		Write-Host "Please take this into consideration when using the Search Results."          -ForegroundColor Yellow
		Write-Host "===========================================================================" -ForegroundColor Yellow
		Read-Host -Prompt "Please press Enter after reading the warning above."
	}
	
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
		Write-host "===================================================="  -ForegroundColor Red
		Write-Host "Take the Search Results above and Investigate." -ForegroundColor Red
		Write-host "===================================================="  -ForegroundColor Red
		ShowNoDeleteMenu
	}
	#If the search was any other type, show the regular Actions menu that allows Delete.
	ShowMenu
}

#Function to clear all of the Vars set by SADPhishes
Function ClearSADPhishesVars {
	Clear-Variable -Name AttachmentName
	Clear-Variable -Name AttachmentNameSelection
	Clear-Variable -Name ContentMatchQuery
	Clear-Variable -Name DangerousEDiscoverySearch
	Clear-Variable -Name DangerousEDiscoverySearchQuitChoice
	Clear-Variable -Name DangerousSearch
	Clear-Variable -Name DateEnd
	Clear-Variable -Name DateRange
	Clear-Variable -Name DateRangeSeparator
	Clear-Variable -Name DateStart
	Clear-Variable -Name EDiscoverySearchMenuChoice
	Clear-Variable -Name EDiscoverySearchName
	Clear-Variable -Name ExchangeLocation
	Clear-Variable -Name ExchangeSearchLocation
	Clear-Variable -Name mailboxes
	Clear-Variable -Name MailboxSearch
	Clear-Variable -Name MailboxSearches
	Clear-Variable -Name MenuChoice
	Clear-Variable -Name NoDeleteMenuChoice
	Clear-Variable -Name PurgeName
	Clear-Variable -Name PurgeSuffix
	Clear-Variable -Name SearchName
	Clear-Variable -Name SearchType
	Clear-Variable -Name Sender
	Clear-Variable -Name Subject
	Clear-Variable -Name ThisEDiscoverySearch
	Clear-Variable -Name ThisEDiscoverySearchName
	Clear-Variable -Name ThisEDiscoverySearchRun
	Clear-Variable -Name ThisPurge
	Clear-Variable -Name ThisSearch
	Clear-Variable -Name ThisSearchResults
	Clear-Variable -Name ThisSearchResultsLine
	Clear-Variable -Name ThisSearchResultsLines
	Clear-Variable -Name TimeStamp
}

Function PrintSADPhishesVars {
	Write-Host AttachmentName [$AttachmentName]
	Write-Host AttachmentNameSelection [$AttachmentNameSelection]
	Write-Host ContentMatchQuery [$ContentMatchQuery]
	Write-Host DangerousEDiscoverySearch [$DangerousEDiscoverySearch]
	Write-Host DangerousEDiscoverySearchQuitChoice [$DangerousEDiscoverySearchQuitChoice]
	Write-Host DangerousSearch [$DangerousSearch]
	Write-Host DateEnd [$DateEnd]
	Write-Host DateRange [$DateRange]
	Write-Host DateRangeSeparator [$DateRangeSeparator]
	Write-Host DateStart [$DateStart]
	Write-Host EDiscoverySearchMenuChoice [$EDiscoverySearchMenuChoice]
	Write-Host EDiscoverySearchName [$EDiscoverySearchName]
	Write-Host ExchangeLocation [$ExchangeLocation]
	Write-Host ExchangeSearchLocation [$ExchangeSearchLocation]
	Write-Host mailboxes [$mailboxes]
	Write-Host MailboxSearch [$MailboxSearch]
	Write-Host MailboxSearches [$MailboxSearches]
	Write-Host MenuChoice [$MenuChoice]
	Write-Host NoDeleteMenuChoice [$NoDeleteMenuChoice]
	Write-Host PurgeName [$PurgeName]
	Write-Host PurgeSuffix [$PurgeSuffix]
	Write-Host SearchName [$SearchName]
	Write-Host SearchType [$SearchType]
	Write-Host Sender [$Sender]
	Write-Host Subject [$Subject]
	Write-Host ThisEDiscoverySearch [$ThisEDiscoverySearch]
	Write-Host ThisEDiscoverySearchName [$ThisEDiscoverySearchName]
	Write-Host ThisEDiscoverySearchRun [$ThisEDiscoverySearchRun]
	Write-Host ThisPurge [$ThisPurge]
	Write-Host ThisSearch [$ThisSearch]
	Write-Host ThisSearchResults [$ThisSearchResults]
	Write-Host ThisSearchResultsLine [$ThisSearchResultsLine]
	Write-Host ThisSearchResultsLines [$ThisSearchResultsLines]
	Write-Host TimeStamp [$TimeStamp]
}

#Drop the user into the DisplayBanner function (and then Search Type Menu) to begin the process.
DisplayBanner