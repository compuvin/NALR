# NALR
Network Access Level Review

Collects permissions information on a given folder and their subfolders and writes the collected information to an MS Access database. It then reports on permissions changes such as a user or group gaining access to a folder. It does not look at individual file permissions within the folder.

List_Sec_Folder_v3.vbs is the main file. The original credit for this script goes to Alexandre LF coppied from Microsoft Technet. Database and change management code was added to his original script.

Usage: Change the information at the top of the script and then scroll down to the bottom and configure the SendMail function.

DumpGroupMembership.ps1 will dump AD group members to a csv file and run ImportGroups.vbs to import the groups into the database. This can then be used to report on the permissions that an individual user has to the folder.
