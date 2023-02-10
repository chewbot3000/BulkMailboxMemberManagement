<#
.SYNOPSIS
Bulk Manage Mailbox Members

.DESCRIPTION
This script will allow you to grant users mailbox permissions including SendAs permissons. It pulls information
from a CSV named MailboxMembers.csv that needs to have the following column headers: Mailbox, User

In the mailbox column you will need to use the mailbox Alias or UPN minus the @cmh.edu. 
In the user column, you will use their Alias or UPN, also without the @cmh.edu

To prevent the mailbox from automatically mapping to their Outlook (for instance they need it setup as a separate
profile) use the -DisableAutoMap switch.

To include SendAs permissons for the mailbox, use the -AddSendAs Switch

If you need to only add SendAs permissions, use the -SendAsOnly Switch

You can combine switches as well. For example, someone needing mailbox access with SendAs permissions but does not
need the mailbox to auto map, you would combine the switches: i.e. -AddSendAs -DisableAutoMap


.NOTES
When changing data in the CSV, make sure to highlight the cells, right click, and select Delete. (shift cells either
up or to the left). Otherwise, the data will still reside even though it looks blank.


.PARAMETER DisableAutoMap
To prevent the mailbox from automatically mapping to their Outlook (for instance they need it setup as a separate
profile)

.PARAMETER AddSendAs
To include SendAs permissons for the mailbox

.PARAMETER SendAsOnly
If you need to only add SendAs permissions


.EXAMPLE
For a normal ticket, after you have entered data into the CSV and saved it as MailboxMembers.csv, run the following:
PS C:\>.\Add-MailboxMembers.ps1 -AddSendAs

.EXAMPLE
Some requests will be only for SendAs permissions, run the following:
PS C:\>.\Add-MailboxMembers.ps1 -SendAsOnly

.EXAMPLE
For requests where they're asking for mailbox permissions but need to set it up as as secondary Outlook profile:
PS C:\>.\Add-MailboxMembers.ps1 -DisableAutoMap

.EXAMPLE
If needed, switches can be combined:
PS C:\>.\Add-MailboxMembers.ps1 -AddSendAs -DisableAutoMap
#>

#=================================================================================================

#							Script:		Add-MailboxMembers
#							Author:		Zabolyx, Chewie
#							Version:	0.1.5 (Template ver. 0.1.3)
#							Date: 		02/10/2023

#	Changelog at the bottom of the script
#=================================================================================================

#=================================================================================================
#	Parameters
#=================================================================================================
	Param(
		[Switch] $DisableAutomap,
		[Switch] $AddSendAs,
		[Switch] $SendAsOnly
		
	)
	

#=================================================================================================
#	Main Code
#=================================================================================================

#region
	
	IF (Test-Path ".\MailboxMembers.csv"){
		
#		Load the list into an aray
		$aData = Import-CSV -path ".\MailboxMembers.csv"
	}

#		Creates the CSV with headers and quits the script	
	Else{
		Add-Content ".\MailboxMembers.csv" "Mailbox,User"
		
		Write-host "CSV not found in directory. Created MailboxMembers.csv in the same folder as script." -Foregroundcolor Red
		Write-host "Please modify the CSV with the required information and run again." -Foregroundcolor Red
		Write-Host "The script will now exit." -Foregroundcolor Red
		
		Exit
	}

#	Process the list of Users	
	Foreach ($eUser in $aData){
	
		# Checks users account for a mailbox or if account is a Security Group
		If ((($RecipientType=(Get-Recipient $($eUser.User)).RecipientType) -ne "UserMailbox") -and ($RecipientType -ne "MailUniversalSecurityGroup") ) {
			Write-Host $RecipientType
			Write-Host "$($eUser.User) doesn't have a mailbox to add the permission" -ForegroundColor Red
			Continue
		}
		# Checks to see if the user account is enabled
		If (($RecipientType -eq "UserMailbox") -and ($(Get-ADUser $($eUser.User)).Enabled -eq $False)) {
			Write-Host "$($eUser.User) account is disabled" -ForegroundColor Red
			Continue
		}
		# Checks if the account to be granted permissions to has a mailbox
		If (![bool]$(Get-Mailbox $($eUser.Mailbox) -ErrorAction SilentlyContinue)) {
			Write-Host "$($eUser.Mailbox) does not have a mailbox to add permissions to." -ForegroundColor Red
			Continue
		}	
	
		$MailboxID=(get-mailbox "$($eUser.Mailbox)").exchangeGUID
		
		Try{
			IF ($SendAsOnly -eq $False){
					
				IF ($DisableAutomap){
					Add-MailboxPermission -Identity $MailboxID -User "$($eUser.User)" -AccessRights FullAccess -InheritanceType All -Automapping $False
				}
				Else{
					Add-MailboxPermission -Identity $MailboxID -User "$($eUser.User)" -AccessRights FullAccess -InheritanceType All
					
				}
			
				Write-host "Complete - added $($eUser.User) to mailbox $($eUser.Mailbox)" -Foregroundcolor Green
			
			}
		}
			
		Catch{
		
				Write-host "Unforseen Error" -Foregroundcolor Red
				
		}
		
		IF ($AddSendAs -or $SendAsOnly){
			
			Try {
				Add-RecipientPermission -Identity $MailboxID -Trustee "$($eUser.User)" -AccessRights 'SendAs' -Confirm:$False
			
				Write-host "Complete - added $($eUser.User) to mailbox $($eUser.Mailbox) SendAs Permission" -Foregroundcolor Green

			}
			Catch{
			
				Write-host "Unforseen Error" -Foregroundcolor Red
					
			}			
		
		}
		

	}

#endregion
#=================================================================================================
#	Change Log
#=================================================================================================

#	Ver 0.1.0: 09/23/2022 ADD - Initial release
#	Ver 0.1.1: 09/29/2022 FIX - Username not properly being polled
#						  ADD - Colors to write-host
#	Ver 0.1.2: 10/27/2022 FIX - Errors not displaying on screen, just skipping line if there's an
#								error.
#	Ver 0.1.3: 01/13/2023 ADD - Cheks for mailbox and user accounts
#	Ver 0.1.4: 01/20/2023 Fix - Fixed check for RecipientType to include User and Mail Enabled 
#								Security Groups
#	Ver 0.1.5: 02/10/2023 ADD - Switch to SendAsOnly
#
