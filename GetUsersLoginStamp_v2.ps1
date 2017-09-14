# Created by: Dan Lundqvist
# Thanks to: Hey Scripting Guy (https://blogs.technet.microsoft.com/heyscriptingguy/2012/12/11/use-powershell-to-add-two-pieces-of-csv-data-together/)
# 
# Scenario:
# We wanted all the users that has enabled account and get their timestamps from AD and Exchange
# Then if they haven't logged in for X days ($Datestr). Disable AD Account and move to DISABLED-OU
# If they have exchange that haven't been accessed for X days($Datestr), delete mailbox.
# All OK users, move to a seperate AD Group to keep track of them.
#
# Normal output (Database Column is used to verify if a user has mailbox or not, either they logged in or not)
# Spaces are added only here, to see all text easier
# Name     , SamAccountName , LastLogonDate    , LastLogoffTime   , LastLogonTime    , Database
# Jane Doe , jane.doe       , 2017-07-03 12:00 , 2017-09-03 12:00 , 2017-09-03 12:00 , DB1
# John Doe , john.doe       , 2017-07-03 12:00 ,                  ,                  , DB1
#  
# 3500 user takes about 5 minutes to run
# Change OU's in $OUs 
# Change <servername> in $Session configuration

Import-module ActiveDirectory
#Imports the session to Exchange server
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<ServerName>/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session

$Logroot = "C:\Temp\UserCheck"                     #Root folder
$LogdirDump = "$Logroot\CSV_Dump.txt"              #Dump file
$LogdirADSAM = "$Logroot\Users_ADSAM.csv"          #File location for all AD users
$LogdirAdDisable = "$Logroot\Users_Ad_Disable.csv" #Users that are NOT OK
$LogdirExDisable = "$Logroot\Users_Ex_Disable.csv" #Users that are NOT OK
$LogdirInfo = "$Logroot\Users_Informationen.csv"   #Final file location

$DateStr = (Get-Date).adddays(-365).ToString("yyyy-MM-dd") #Sets the date to check with
$ADGroup = "Sharepoint_AD_Users"

#Remove files
Remove-Item -Recurse "$Logroot\*"

#Add content to top of csv file
"Name,SamAccountName,LastLogonDate,LastLogoffTime,LastLogonTime,Database" | Out-File $LogdirDump
"SamAccountName" | Out-File $LogdirAdDisable
"SamAccountName" | Out-File $LogdirExDisable

#OUs to verify against 
$OUs = "OU=Employees,OU=Users,OU=Accounts,OU=Company,DC=domain,DC=local"
$DisabledOU = "OU=Disabled,OU=Accounts,OU=Company,DC=domain,DC=local"

#Gets all users that active
ForEach ($OU in $OUs)
{Get-ADUser -SearchBase $OU -Filter {Enabled -eq $true} | Select SamAccountName | Export-CSV $LogdirADSAM -Encoding UTF8 -NoTypeInformation -Delimiter ";" -Append}

Import-Csv -Path $LogdirADSAM | ForEach-Object {
#Resets the value of variables each time the loop restarts
$AdInfo = $ExInfo = $ADandEX = $null
#Gets information from AD: Name, EmployeeID, Lastlogondate
$ADInfo = Get-ADUser -Identity $_.SamAccountName -Properties LastLogonDate | Select Name,SamAccountName,LastLogonDate | ConvertTo-Csv -NoTypeInformation
#Gets information from Exchange Lastlogofftime, lastlogontime, database
$ExInfo = Get-Mailboxstatistics -Identity $_.SamAccountName -EA SilentlyContinue | Select LastLogoffTime,LastLogonTime,Database | ConvertTo-Csv -NoTypeInformation

#If user dont have an exchange account then skip the output from exchange and only takes AD
If ($ExInfo -eq $null) 
        {
            $ADandEX += "{0}" -f $ADInfo[1]
        }
#If user has both exchange and AD, then output that to the file
Else
        {
            $ADandEX += "{0},{1}" -f $ADInfo[1],$ExInfo[1]
        }
        #Adds the information to file
        $ADandEX | Out-File $LogdirDump -Append 
}
#Removes the session from Exchange server
Remove-PSSession $Session

#This converts the txt file to a normal csv file!
Get-Content $LogdirDump | ConvertFrom-Csv | Export-Csv $LogDirInfo -NoTypeInformation -Encoding UTF8 

#Gets all the account needed to be Disabled!
Import-Csv -Path $LogdirInfo | ForEach-Object {
    If (($_.LastLogonDate -eq $null) -or ([string]::IsNullOrEmpty($_.LastLogonDate)) -or ([string]::IsNullOrWhiteSpace($_.LastLogonDate))) {$AdStamp = "1000-01-01"}
    Else {$AdStamp = (Get-Date $_.LastLogonDate).ToString("yyyy-MM-dd")}
        If (($AdStamp -lt $DateStr))
        {
            $_.SamAccountName | Out-File $LogdirAdDisable -Append
        }
        Else {}
}

#Gets all the mailboxes needed to be Disabled!
Import-Csv -Path $LogdirInfo | ForEach-Object {
    If ($_.LastLogonTime -eq $null -or [string]::IsNullOrEmpty($_.LastLogonTime) -or [string]::IsNullOrWhiteSpace($_.LastLogonTime)) {$ExStamp = "1000-01-01"}
    Else {$ExStamp = (Get-Date $_.LastLogonTime).ToString("yyyy-MM-dd")}
        If (($ExStamp -lt $DateStr) -and (![string]::IsNullOrEmpty($_.Database)) -and (![string]::IsNullOrWhiteSpace($_.Database)))
            {
                $_.SamAccountName | Out-File $LogdirExDisable -Append
            }
        Else {}
    }

Remove-Item -Recurse "$Logroot\*" -Exclude Users_Informationen.csv,Users_Ad_Disable.csv,Users_Ex_Disable.csv

############# DISBALES USERS AND ADD DESCRIPTION & INFO VALUE #############
#Disable the user added to $LogdirAdDisable and move to seperate OU ($DisableOU). Also add an info value, this is to check is they have been disabled for x days. Then DELETE THEM!
#Import-Csv -Path $LogdirAdDisable | ForEach-Object {
#$TimeStamp = Get-Date -Format g
#$UserDN = (Get-ADUser -Identity $_.SamAccountName).distinguishedName
#Move-ADObject -Identity $UserDN -TargetPath $DisabledOU
#Set-ADUser -Identity $_.SamAccountName -Enabled $False -Description "User disabled: $TimeStamp, By HelpDesk" -Replace @{info="$TimeStamp"}
#}
##########################################################################

########### CREATES AN AD GROUP FOR SHAREPOINT USE #######################
#First Clears the group
#Get-ADGroupMember $ADGroup | ForEach-Object {Remove-ADGroupMember $ADGroup $_ -Confirm:$false}

#Then adds all the new members to this group
#Import-CSV $LogdirAdOK | % { Add-ADGroupMember -Identity $ADGroup -Members $_.SamaccountName }
##########################################################################

############ INFO TAG ALL DISABLED USERS!! ###############
#$DateUser = (Get-Date).adddays(-120).ToString("yyyy-MM-dd") #Sets the date to check with
#ForEach ($OU in $OUs)
#{Get-ADUser -SearchBase $OU | Select SamAccountName | Export-CSV $LogdirSAM -Encoding UTF8 -NoTypeInformation -Delimiter ";" -Append}

#Import-Csv -Path $LogdirSAM | ForEach-Object {Set-ADUser -Identity $_.SamAccountName -Enabled $False -Replace @{info="$TimeStamp"}}

#### USED TO CHECK IF INFO IS OLD ######
#Import-Csv -Path $LogdirSAM | ForEach-Object {
#$UserInfoAttribute = Get-ADUser -Identity $_.SamAccountName -Properties info | Select info
#If ($UserInfoAttribute.info -lt $DateUser) 
#{
#    Remove-ADUser -Identity $_.SamAccountName
#}
#Else {}
#}
#######################################################