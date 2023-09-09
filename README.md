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
  ### Main Screen
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266756779-a91e5b2b-6e5a-4e02-b002-65c3d3ec290c.png">
</p>

  ### User Screen
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266756787-921339e4-9631-41c2-876e-6965eb67a333.png">
</p>

  ### Groups Screen
<p align="center">
  <img  src="https://user-images.githubusercontent.com/30836537/266756764-a7a9be6b-4482-45cf-a653-eed7294d7ec2.png">
</p>


