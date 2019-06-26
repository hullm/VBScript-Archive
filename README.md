# Active Directory Scripts
### AD Report 1.02.vbs
Scans an Organizational Unit and reports information about the users in that OU. It will output the data to a CSV file. You can use this to check user properties such as home folder, script, and profile path information. Then it will check the username format for errors, verify that the user has a home folder and report the size of the home folders. The script is dynamic in what it reports, for example if the users don't have a setting for home folders it won't check to see if a folder exists or try to report the size. It will also attempt to determine your username format.

### Count Exchange Accounts 1.0.vbs
This script will count the number of mailboxes on your exchange server. It pulls the server name from the first account with a mailbox in the Domain Admins group.

### Display FSMO Roles 1.0.vbs
This will query Active Directory for the name of the FSMO role holders and display them to the user. PDC Emulator, RID Master, Schema Master, Infrastructure Master, Domain Naming Master.

### Fix Display Names 1.02.vbs
Loops through an Organizational Unit and changes the displayname and CN attributes to lastname, firstname format. It looks at the display name for it's information. In order to run this properly the display name has to be in firstname lastname format. If it detects a comma in the display name it will skip that user. It also can run in a report mode so you can see what it would do to the users before it runs. Use caution when running this script it has potential to mess things up big time.

Note: ADHelper.wsc required

### Fix Printers 1.0.vbs
This script will help you fix network printers after a server has been renamed.

### Import Phone Numbers into AD 1.0.vbs
This script will read a CSV file in the format lastname,firstname,phone number and try to locate the user in Active Directory and if found it will add the phone number. This script is lacking in error detection, use at your own risk.

Note: ADHelper.wsc required

### Import Users 1.0.vbs
Imports users into Active Directory from a CSV file with the users listed in lastname, firstname format. The script will generate passwords, make home folders, and add users to groups. You can run the script in a report only mode to verify the users before they are entered. You should use Clean CSV File 1.0.vbs to prepare an import file before you run the script. If there are errors in the import file the script will not run.

Note: ADHelper.wsc required

### Last Logon Report - Computer.vbs - Written by Richard L. Mueller
This script was not written by me but I did modify it a little. It will now run without needing to be edited. It will scan all your Domain Controllers and report when each computer account last logged onto the domain. This can be useful when tracking down orphaned computer accounts.

### Last Logon Report - User.vbs - Written by Richard L. Mueller
This script was not written by me but I did modify it a little. It will now run without needing to be edited. It will scan all your Domain Controllers and report when each user account last logged onto the domain. This can be useful when tracking down orphaned user accounts.

### Modify Logon Hours 1.0.vbs
This script will loop through each user in an OU and set their logon hours setting using the Net User command.

Note: ADHelper.wsc required

### Modify TS Settings 1.0.vbs
This script uses a program called TSCMD from a company called System Tools, to modify all the users terminal services settings in an OU.

Note: ADHelper.wsc required

### OU Password Reset 1.01.vbs
Use this to reset everyone's password in an Organizational Unit. You can assign everyone the same password or have the script generate a unique password for each user. The script will create a log file with the passwords.

Note: ADHelper.wsc required

### Remove Users Without Folders 1.0.vbs
This script will delete any user that doesn't have a home folder. The idea is that you would have run the script "Delete Empty Folders 1.0.vbs" first. This way anyone who has never used their account can be removed. USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE USER ACCOUNTS!!!!

Note: ADHelper.wsc required

### Replicate Two DC's 1.0.vbs
This will replicate two domain controllers. When you run the script it will prompt you for the name of the source and destination servers.

### Set Mailbox Limits 1.0.vbs
Loops through each user in an Organizational Unit and sets their mailbox limit settings. You can also use it to clear or disable the settings.

Note: ADHelper.wsc required

### Swap First and Last Names 1.0.vbs
Swaps the first and last name property of each user in an OU. This was written to fix a problem caused by a bug in Fix Display Names 1.02.vbs (The bug has been fixed)

Note: ADHelper.wsc required

### SWS Settings 1.0.vbs
Scans the domain for information needed to switch Symantec Web Security to LDAP mode and displays the information to the user. This makes it easier to set up SWS.

### User Mod Add 1.0.vbs
This script will loop thru each user in an OU and set certain properties. The properties it can set are Description, Home Directory, Home Drive, Profile Path and Script.

Note: ADHelper.wsc required

### User Mod Delete 1.0.vbs
This script can be used to clear certain properties in all users in an OU. The properties it can clear are Description, Home Directory, Home Drive, Profile Path and Script.

Note: ADHelper.wsc required

# Desktop Scripts
### AutoLogon 1.0.vbs
This script will set a user to auto logon to a computer. It modifies the appropriate registry setting to do this. This was written for an elementary where they wanted all the computers to log in with a generic elem account.

### Change Background Color 1.0.vbs
This script was written for a school that wanted to change the background color on all the PC's in a lab. They were going to need to change the background colors often so this script was created to allow the teacher to easily do this. She would change the settings in the script on the server and when the students logged in the background color changed.

### Change Script Type 1.0.vbs
This is a small simple script that will change your script engine to cscript or wscript. I wrote this when I was working with a script that needed cscript and most of my others use wscript. It allowed me to quickly switch between them.

### Display Computer Name 1.0.vbs
Shows the name of your computer. We have users run this from a web server so they can easily get their computer name when we do remote support.

### Display Logon Server 1.0.vbs
Use this to show the logon server and the currently logged in user. I wrote it to learn how to deal with environmental variables.

### Fix Sync Warnings 1.0.vbs
In Windows 2k/XP with offline files enabled the computer may pause with a sync warning. This script will change a registry setting so the warning will still show but the computer will continue to shut down.

### Force Classic Start Menu 1.0.vbs
This was used to force a Windows 2003 terminal server to have all users see the classic start menu.

### Guessing Game 1.0.vbs
Just for fun. This was an early script that tested a lot of the skills I had been learning. There is no action in this game. It is a little boring.

### Map Network Drives 1.0.vbs
Maps network drives on a workstation. If it fails it logs a warning in the event viewer. This is a derivative of the Printer Install 2.0 script (see below)

### Mod Proxy 1.0.vbs
Modifies the proxy settings on a computer. This was created for a school that was having a problem with Group Policies. It should work with Windows 9x/Me as well although I haven't tested it.

### Network Drives 1.01.vbs
This will display all the network drives on your computer with a description. You can use it to document your network drives. The user then can run the script and see what network drives they have and what they are for.

### Outlook 2003 Fix 1.0.vbs
By default Outlook 2003 blocks certain types of file attachments. You can use this script to allow some of them in.

### Outlook XP Fix 1.0.vbs
By default Outlook XP blocks certain types of file attachments. You can use this script to allow some of them in.

### Password Generator 1.0.vbs
Just a small script that will generate a list of passwords.

### Printer Install 2.0.vbs
Installs network printers on a workstation. If it fails it logs a warning in the event viewer.

### Printer Uninstall 2.0.vbs
Uninstalls network printers from a workstation. If it fails it logs a warning in the event viewer.

### WSH Version 1.0.vbs
Displays the version of Windows Scripting Host on a computer. I wrote this one to find out what version was installed on different Windows 95 computers.

# Folder Scripts
### Archive Students Data 1.0.vbs
This script will read in a CSV and compare the list of users to users in an OU. If the user isn't in the CSV then their data is moved to an archive folder and their account is deleted. If you have an OU that has users in it this can be used to remove users that aren't needed anymore. The CSV would contain the users who are supposed to be in the OU in a lastname, firstname format USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE USER ACCOUNTS!!!!

Note: ADHelper.wsc required

### Clean Profiles 1.0.vbs
The will clear all the users profiles on a computer except for some of the key ones. It will then report the amount of space that was freed. This has been used in computer labs at the end of the year to clean the computer or to troubleshoot group policy problems.

### Clean Remote Profiles 1.0.vbs
Same as the Clean Profiles script except it will clean all computers in an OU. It uses DelProf from Microsoft to remove the profiles. It will also create a log with the computers it cannot contact.

### Compact MDB.vbs - Written by Danny Lesandrini
Compacts an access database. The computer you run it on needs to have access installed.

### Copy Favorites Logoff 1.01.vbs
This can be used as a solution to the lack of Favorites redirection in Active Directory. When the user logs off their favorites are copied from the local profile to their home folder. Use this with Copy Favorites Logon 1.01

### Copy Favorites Logon 1.01.vbs
This can be used as a solution to the lack of Favorites redirection in Active Directory. When the user logs on their favorites are copied from their home folder to their local profile. Use this with Copy Favorites Logoff 1.01

### Copy Files 1.0.vbs
All this does is copy files and folders from one place to another.

### Delete Empty Folders 1.0.vbs
This script will delete all the empty subfolders in a folder. This is used to remove empty student folders. USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE DATA!!!!

### Folder Size 1.0.vbs
This will report the size of all subfolders in a folder.

### Home Folders Check 1.0.vbs
You can use this script to verify that users in an OU have home folders.

### Make Folder From OU 1.01.vbs
Creates a folder for each user in an OU using the users username as the folder name. Then it sets up the permissions on each folder.

### Make Folders From Group 1.01.vbs
Creates a folder for each member of an group using the users username as the folder name. Then it sets up the permissions on each folder.

### Modify Folder Ownership 1.01.vbs
This script is designed to use the SubinACL program found in the Windows 2000 Resource Kit to reset Ownership on users home folders. It will assign ownership of all files and folders in each users folder to the proper account.

### Modify Permissions 1.0.vbs
This script can be used to reset permissions on users home folders. It can take the folder name and use it as a username and assign permissions accordingly.

### Move Home Folders 1.01.vbs
This script just moves home folders from one location to another. It will update the users home folder settings in there profile and fix the permissions on the new folders. This proves to be very useful when moving home folders from one server to another.

Note: ADHelper.wsc required

### Move Unwanted Files 1.0.vbs
This script can be used to scan student home folders and remove files you don't want them to have. It will move the files to another directory then log where they came from. You can also have it replace the file with a text file stating that the file was removed.

### Orphaned Folder Check 1.1.vbs
Loops through home folders and verifies the folder has a user associated with it by querying Active Directory. You can use Orphaned Folder Check 1.0.vbs to have it verify against users in one OU. This one can be useful if you have multiple OU's and users are sometimes moved between them.

### Set Quota for Users in OU 1.0.vbs
This script will set a quota limit for everyone in an OU. NOTE: This version works but can use some improvements, like an exception list and it could easily be made to run on multiple OU's at once. Maybe with the next version.

Note: ADHelper.wsc required
