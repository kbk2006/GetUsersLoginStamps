# Created by: Dan Lundqvist
# Thanks to: Hey Scripting Guy (https://blogs.technet.microsoft.com/heyscriptingguy/2012/12/11/use-powershell-to-add-two-pieces-of-csv-data-together/)
# 
# Scenario:
# We wanted all the users that has enabled account and get their timestamps from AD and Exchange
#
# Normal output:
# Name     , SamAccountName , LastLogonDate    , LastLogoffTime   , LastLogonTime
# Jane Doe , jane.doe       , 2017-07-03 12:00 , 2017-09-03 12:00 , 2017-09-03 12:00
#
# 3500 user takes about 5 minutes to run

Import-module ActiveDirectory
#Imports the session to Exchange server
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<servername>/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session

$Logroot = "C:\Temp\UserCheck2"                  #Root folder
$LogdirADSAM = "$Logroot\Users_ADSAM.csv"        #File location for all AD users
$LogdirDump = "$Logroot\CSV_Dump.txt"            #Dump file
$LogdirInfo = "$Logroot\Users_Informationen.csv" #Final file location

#Remove all files
Remove-Item -Recurse "$Logroot\*"

#Add content to top of csv file
"Name,SamAccountName,LastLogonDate,LastLogoffTime,LastLogonTime" | Out-File $LogdirDump
#OUs to verify against #$DateStr = (Get-Date).adddays(-365).ToString("MM-dd-yyyy")
$OUs = "OU=Employees,OU=Users,OU=Accounts,OU=Company,DC=domain,DC=local","OU=Employees,OU=Users,OU=Accounts,OU=Company,DC=domain,DC=local","OU=Disabled,OU=Accounts,OU=Company,DC=domain,DC=local"

#Gets all users that active from above OU's
ForEach ($OU in $OUs)
{Get-ADUser -SearchBase $OU -Filter {Enabled -eq $true } | Select SamAccountName | Export-CSV $LogdirADSAM -Encoding UTF8 -NoTypeInformation -Delimiter ";" -Append}

Import-Csv -Path $LogdirADSAM | ForEach-Object {
#Resets the value of variables each time the loop restarts
$AdInfo = $ExInfo = $ADandEX = $null
#Gets information from AD: Name, EmployeeID, Lastlogondate
$ADInfo = Get-ADUser -Identity $_.SamAccountName -Properties LastLogonDate | Select Name,SamAccountName,LastLogonDate | ConvertTo-Csv -NoTypeInformation
#Gets information from Exchange Lastlogofftime, lastlogontime
$ExInfo = Get-Mailboxstatistics -Identity $_.SamAccountName -ErrorAction SilentlyContinue | Select LastLogoffTime,LastLogonTime | ConvertTo-Csv -NoTypeInformation

#If user dont have an exchange account the skip the output from exchange and only takes AD
If ($ExInfo -eq $null) 
        {
            $ADandEX += "{0}" -f $ADInfo[1]
        }
#If user has both exchange and AD, then output that to the file
Else
        {
            $ADandEX += "{0},{1}" -f $ADInfo[1],$ExInfo[1]
        }
        $ADandEX | Out-File $LogdirDump -Append
    }
#Removes the session from Exchange server
Remove-PSSession $Session

#This converts the txt file to a normal csv file!
Get-Content $LogdirDump | ConvertFrom-Csv | Export-Csv $LogDirInfo -NoTypeInformation -Encoding UTF8 | Remove-Item $LogdirDump $LogdirADSAM