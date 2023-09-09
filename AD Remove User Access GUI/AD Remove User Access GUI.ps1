<#
.DESCRIPTION
  Script remove all groups of user  
.NOTES
    NAME:AD Remove User Acess
    AUTHOR: Matheus Andrade
    LASTEDIT: 09/09/2023 

#>

$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
#Inicialização do formulário
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();
#Inicialização do formulário
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='AD Remove User Access GUI'
$main_form.Width = 350
$main_form.Height = 150
$main_form.BackColor = "white"
$main_form.FormBorderStyle = 'Fixed3D'
$main_form.MaximizeBox = $false
$Font = New-Object System.Drawing.Font("Verdana",8)
$main_form.Font = $Font
$text2 = New-Object System.Windows.Forms.label
$text2.Location = New-Object System.Drawing.Size(7,40)
$text2.Size = New-Object System.Drawing.Size(270,15)
$text2.ForeColor = "black"
$text2.Text = "Type the username:"
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(7,80)
$Button.Size = New-Object System.Drawing.Size(150,25)
$Button.Text = "Remove Access Groups"
$TextBoxUsername = New-Object System.Windows.Forms.TextBox
$TextBoxUsername.Location = '135,35'
$TextBoxUsername.Size = '125,50'


$global:username

$diretorio = 'C:\temp'
$Button.Add_Click({ 
    $global:username = $TextBoxUsername.Text
    revogarAcessos
})

Function verificarDiretorio{

if (  (Test-Path -Path $diretorio -PathType Container)) {
    "existe"
}
else {
     New-Item -Path $diretorio -ItemType Directory -ErrorAction Stop | Out-Null #-Force
     "Diretory doesn't exist, Diretory has created!"
}

}

Function backupAcessos {
Get-ADPrincipalGroupMembership $global:username  | Select-Object -ExpandProperty name | Out-File -FilePath $diretorio\$global:username.txt
}

$TextBoxUsername.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
    $global:username = $TextBoxUsername.Text
        revogarAcessos
    }
})

Function revogarAcessos {
$mb = [System.Windows.Forms.MessageBox]
$mbIcon = [System.Windows.Forms.MessageBoxIcon]
$mbBtn = [System.Windows.Forms.MessageBoxButtons]
[System.Windows.Forms.Application]::EnableVisualStyles();
$result = $mb::Show("Are you sure you want to remove the user's access " +$global:username+"?", "AD Remove User Access GUI", $mbBtn::YesNo, $mbIcon::Exclamation)
Echo $result
If ($result -eq "Yes") {
    verificarDiretorio
    backupAcessos

    Get-ADUser -Identity $global:username -Properties MemberOf | ForEach-Object {
  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false

  $TextBoxUsername.Text = ""

}
    	
[System.Windows.Forms.MessageBox]::Show("User access"+$global:username+" has been revoked!, Backup created at C://temp2","Remover Acessos",0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)


}
Else {

$TextBoxUsername.Text = ""
}
   
}

$main_form.Controls.Add($TextBoxUsername)
$main_form.Controls.Add($Button)
$main_form.Controls.Add($text2)
$main_form.ShowDialog()
