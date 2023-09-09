<#
.DESCRIPTION
  Script to remove user acess by a list 
.NOTES
    NAME: AD Remove User Acess List
    AUTHOR: Matheus Andrade
    LASTEDIT: 09/09/2023 

#>
$result = @()
#Need to fill .txt with Samaccountname in each line 
 Get-Content C:\temp\userList.txt |
 ForEach-Object{

 Get-ADUser -Identity $_ -Properties MemberOf | ForEach-Object {
  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false


  }

 }