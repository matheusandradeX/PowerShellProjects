<#
.SYNOPSIS
  Script to show date and time from sytem 
.DESCRIPTION
 Script to show date and time from sytem 
.NOTES
  Version:        1.0
  Author:         Matheus Andrade
  Creation Date:  2023
#>
mode con:cols=70 lines=8
function Show-Menu
{
    param (
        [string]$Title = 'My Menu'
    )
    Clear-Host
    Write-Host ""
    
    Write-Host "1: Show Date"
    Write-Host "0: Exit"
  
}

do
 {
     Show-Menu
     $selection = Read-Host "Choice option:"
     cls
     switch ($selection)
     {
         '1' {
            Write-Host ""
             date
             
         } 
     }
     pause
 }
 until ($selection -eq '0')