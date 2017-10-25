Param(
	[String]$File
)

##
## Script Created by Anderson Patricio
##

#General Settings
clear

#Validating the File switch
$tPath = Test-Path $File
If ($tpath -eq $False) { 
	Clear
	Write-Host
	Write-Host "Error: File could not be found! Please check you file name." -ForeGroundColor:Red
	Write-Host
	exit 
} Else {
	$vTranscriptPath = (Get-Item -Path ".\" -Verbose).FullName + "\" + $file + ".transcript"
	Start-Transcript -Path $vTranscriptPath -Append
	$users = import-csv $file
}

Import-Module ActiveDirectory


## Listing the users to be migrated
Write-Host
Write-Host "Users to be migrated..." -ForeGroundColor:Yellow
$users


## Generating a list of all send as of the organization
Write-Host
Write-Host ".:. Generating a list of all Send As permissions in the Org that use the current mailboxes" -ForeGroundColor:Yellow
Write-Host "    This Process will take a while, go for a Tim Hortons!"
Write-Host
$ORGSendAS = get-mailbox -ResultSize:Unlimited | Get-ADPermission | ? {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") }

$vCount=0
While ($vCount -le ($users.count)) {
	$ORGSendAS | where-object { $_.User -like ("*" + $users[$vCount].mailbox)} | Select Identity,User
	$vCount++
}

Write-Host
Write-Host ".:. Validating Send-As..." -ForegroundColor:Yellow
Write-Host "    If there is a mailbox listed here, that account should be added to the migration list"
Write-Host
$users | % { get-mailbox $_.mailbox | Get-ADPermission | ? { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} }

Write-Host
Write-Host ".:. Validating Send on Behalf.." -ForeGroundColor:Yellow
Write-Host "    If there is a mailbox listed here, that account should be added to the migration list"
Write-Host
$users | % { get-mailbox $_.mailbox | where { $_.GrantSendOnBehalfTo -ne $null } }

Write-Host
Write-Host ".:. Validating Azure Licenses..." -ForeGroundColor:Yellow
Write-Host
$users | % { If ($_.Profile -eq "gold")	 {if ((Get-ADGroupMember AZLic-gold).SamAccountName -contains $_.Mailbox -eq $True) { Write-Host "License Okay!"} Else { Add-ADGroupMember AZLic-Gold -Member $_.Mailbox; Write-Host "User $_.Mailbox was added to AZLic-Gold" -ForeGroundColor:Cyan} }}
$users | % { If ($_.Profile -eq "silver") {if ((Get-ADGroupMember AZLic-silver).SamAccountName -contains $_.Mailbox -eq $True) { Write-Host "License Okay!"} Else { Add-ADGroupMember AZLic-Silver -Member $_.Mailbox; Write-Host "User $_.Mailbox was added to AZLic-silver" -ForeGroundColor:Cyan} }}
$users | % { If ($_.Profile -eq "bronze") {if ((Get-ADGroupMember AZLic-bronze).SamAccountName -contains $_.Mailbox -eq $True) { Write-Host "License Okay!"} Else { Add-ADGroupMember AZLic-Bronze -Member $_.Mailbox; Write-Host "User $_.Mailbox was added to AZLic-Bronze" -ForeGroundColor:Cyan} }}


Write-Host
Write-Host ".:. Validating On-Prem Archive mailboxes... " -ForeGroundColor:Yellow
Write-Host
$users | % { if ( (get-mailbox $_.mailbox).archiveDatabase -ne $null ) { if ($_.profile -ne 'gold') { write-host $_.Mailbox ": This mailbox has an Archive Database, however it does not have gold associated to it, please check the license" } } }

Stop-Transcript