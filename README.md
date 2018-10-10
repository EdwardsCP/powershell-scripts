# Powershell Scripts
Just a place where I'm making random powershell scripts available for use.

BitLocker-Logging.ps1 is a quick tool to write BitLocker Status to a log file

SADPhishes.ps1 is a tool I'm working on to use Exchange 2016's Compliance Search to delete emails that have been identified as phishing


# SADPhishes.ps1 Flowchart
![SADPhishesFlowchart](/SADPhishes%20Screenshots/SADPhishesFlowchart.jpg)

# SADPhishes.ps1 Basic Usage Screenshots
The screenshots below show an example of a workflow that can be taken through the SADPhishes script.  In this example...

...the user created a Compliance Search query using a Subject, Sender Address, and Date Range

...the user did not include an Attachment in the search

...the user searched against one 1 mailbox (instead of All mailboxes, or a group)

...the search returned 1 item

...the user chose to create and execute an In-Place eDiscovery Search based on the results of the Compliance Search

...the user launched the eDiscovery Search Results in their browser

...the user Deleted all of the items that were returned by the Compliance Search

...the user was returned to the top level Search Options Menu


![SADPhishes1](/SADPhishes%20Screenshots/SADPhishes1.png)

![SADPhishes2](/SADPhishes%20Screenshots/SADPhishes2.png)

![SADPhishes3](/SADPhishes%20Screenshots/SADPhishes3.png)

![SADPhishes4](/SADPhishes%20Screenshots/SADPhishes4.png)

![SADPhishes5](/SADPhishes%20Screenshots/SADPhishes5.png)

![SADPhishes6](/SADPhishes%20Screenshots/SADPhishes6.png)
