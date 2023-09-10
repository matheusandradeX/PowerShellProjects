# PowerShellProjects


## AD User Consult

 A simple tool to consult user attributes using like Active Diretory to Service Desk team
## Dependencies
Needs install Remote Server Administration Tools (RSAT)  that code on your powershell on Windows Client, Widows Server doesn't require that installation 

```
Add-WindowsCapability –online –Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

 ### Features

 - Search user object from OU'S by display name and show a list
 - Show attributes from selected user
- Show Group Access from selected user
- Copy user data to cliboard (userfull to service desk team paste on ITSM Incident Ticket

### Screenshots 
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266809189-104e4f37-d79c-4ce5-91e1-e6e52d355cae.png">
</p>

## AD Import User from CSV

A tool to import a list of users in your Active Diretory by CSV file 
## Dependencies
Needs to change UPN value in the script  to your domain name and fill CSV model File with user data

### Screenshots 
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266808831-37cd67f2-2359-46ad-a250-e0810bfc46db.png">
</p>

## AD Reset Password
A tool to reset password by sammaccountname 
 ### Features

 - Select doamin controller 
 - Show user info
-  Unlock user
-  Generate strong random password
  ### Screenshots 
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266810720-eba5f5eb-6aee-4979-bf73-951d8b57fbe4.png">
</p>




