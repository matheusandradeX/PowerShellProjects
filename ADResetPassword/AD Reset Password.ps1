<#
.DESCRIPTION
  Script for password reset.
  Need to change Server Name to ours 
.NOTES
    NAME: AD Remove User Acess
    AUTHOR: Matheus Andrade
    LASTEDIT: 09/09/2023 

#>
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName 'System.Web'
Add-Type -AssemblyName PresentationCore
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='AD Reset Password'
$main_form.Width = 650
$main_form.Height = 500
$main_form.FormBorderStyle = 'Fixed3D'
$main_form.MaximizeBox = $false
$Font = New-Object System.Drawing.Font("Verdana",12)
$main_form.Font = $Font
$main_form.StartPosition = "CenterScreen"

$global:nomeDominio

$textDominio = New-Object System.Windows.Forms.label
$textDominio.Location = New-Object System.Drawing.Size(7,10)
$textDominio.Size = New-Object System.Drawing.Size(200,20)
$textDominio.Text = "Select the server: "

$comboBoxDominio = New-Object system.Windows.Forms.ComboBox
$comboBoxDominio.text = “”
$comboBoxDominio.Location = New-Object System.Drawing.Size(219,10)
$comboBoxDominio.width = 202
$comboBoxDominio.autosize = $true

$comboBoxDominio.Items.AddRange(@("ambiente.homologacao", "ambiente.dev", "ambiente.prod"))
$comboBoxDominio.SelectedIndex = 1

$comboBoxDominio.add_SelectedIndexChanged({
     $selectedValue = $comboBoxDominio.SelectedItem.ToString()
     $global:nomeDominio=  $selectedValue
  
})


$panel = New-Object System.Windows.Forms.Panel
$panel.Size = New-Object System.Drawing.Size(600, 350)
$panel.Location = New-Object System.Drawing.Point(15, 75)
$panel.BackColor = "white"
$panel.BorderStyle =1


$text1 = New-Object System.Windows.Forms.label
$text1.Location = New-Object System.Drawing.Size(7,40)
$text1.Size = New-Object System.Drawing.Size(280,25)
$text1.Text = "Type user name:"
$global:username=""

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(430,40)
$Button.Size = New-Object System.Drawing.Size(150,30)
$Button.Text = "Search"

$TextBoxUsuario = New-Object System.Windows.Forms.TextBox
$TextBoxUsuario.Location = '220,40'
$TextBoxUsuario.Size = '200,50'

$textInformacoes1 = New-Object System.Windows.Forms.label
$textInformacoes1.Location = New-Object System.Drawing.Size(7,25)
$textInformacoes1.Size = New-Object System.Drawing.Size(62,20)
$textInformacoes1.Text = "Name:"

$textInformacoesBox1 = New-Object System.Windows.Forms.TextBox
$textInformacoesBox1.Location = '145,25'
$textInformacoesBox1.Size = '300,20'
$textInformacoesBox1.ReadOnly = $true


$textEmail = New-Object System.Windows.Forms.label
$textEmail.Location = New-Object System.Drawing.Size(5,75)
$textEmail.Size = New-Object System.Drawing.Size(69,20)
$textEmail.Text = "E-mail:"

$TextBoxEmail = New-Object System.Windows.Forms.TextBox
$TextBoxEmail.Location = '145,75'
$TextBoxEmail.Size = '300,20'
$TextBoxEmail.ReadOnly = $true

$textDepartamento = New-Object System.Windows.Forms.label
$textDepartamento.Location = New-Object System.Drawing.Size(5,120)
$textDepartamento.Size = New-Object System.Drawing.Size(140,20)
$textDepartamento.Text = "Department:"

$TextBoxDepartamento = New-Object System.Windows.Forms.TextBox
$TextBoxDepartamento.Location = '145,120'
$TextBoxDepartamento.Size = '300,20'
$TextBoxDepartamento.ReadOnly = $true


$textSenha = New-Object System.Windows.Forms.label
$textSenha.Location = New-Object System.Drawing.Size(7,165)
$textSenha.Size = New-Object System.Drawing.Size(135,20)
$textSenha.Text = "New Password:"

$TextBoxSenha = New-Object System.Windows.Forms.TextBox
$TextBoxSenha.Location = '145,165'
$TextBoxSenha.Size = '300,20'
$TextBoxSenha.ReadOnly = $true




$textCheckBox2 = New-Object System.Windows.Forms.label
$textCheckBox2.Location = New-Object System.Drawing.Size(25,235)
$textCheckBox2.Size = New-Object System.Drawing.Size(330,20)
$textCheckBox2.Text = "Must change password at next logon"

$checkbox2 = new-object System.Windows.Forms.checkbox
$checkbox2.Location = new-object System.Drawing.Size(9,230)
$checkbox2.Size = new-object System.Drawing.Size(30,30)
$checkBox2.Enabled = $false;

$textCheckBox1 = New-Object System.Windows.Forms.label
$textCheckBox1.Location = New-Object System.Drawing.Size(25,255)
$textCheckBox1.Size = New-Object System.Drawing.Size(330,20)
$textCheckBox1.Text = "User Blocked"

$checkbox1 = new-object System.Windows.Forms.checkbox
$checkbox1.Location = new-object System.Drawing.Size(9,253)
$checkbox1.Size = new-object System.Drawing.Size(30,30)
$checkBox1.Enabled = $false;


$buttonNovaSenha = New-Object System.Windows.Forms.Button
$buttonNovaSenha.Location = New-Object System.Drawing.Size(10,300)
$buttonNovaSenha.Size = New-Object System.Drawing.Size(150,30)
$buttonNovaSenha.Text = "Password Reset"
$buttonNovaSenha.Enabled = $false

$buttonDesbloquear = New-Object System.Windows.Forms.Button
$buttonDesbloquear.Location = New-Object System.Drawing.Size(430,300)
$buttonDesbloquear.Size = New-Object System.Drawing.Size(150,30)
$buttonDesbloquear.Text = "Unlock"
$buttonDesbloquear.Enabled = $false

 #funcao mudar status se o usuário precisa trocar a senha no proximo login
$checkBox2.Add_Click({
    if (-not $checkBox2.Checked) {
        Set-ADUser -server $global:nomeDominio  -Identity $global:username -ChangePasswordAtLogon $false
    }
     if ( $checkBox2.Checked) {
        [System.Windows.Forms.MessageBox]::Show("User must change password at next logon.")
        Set-ADUser -server $global:nomeDominio -Identity $global:username -ChangePasswordAtLogon $true
    }
})

#funcao para habilitar o botao deve alterar a senha proximo login
function statusProximoLogin {
$checkBox2.Enabled = $true;

 }


#Obtendo a data 
$horaAtual = Get-date -format 'dd/MM/yyy hh:mm:ss tt'
#Obtendo o usuário logado na maquina
$usuarioAtual = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


$Button.Add_Click({ 
    buscar
})

$buttonNovaSenha.Add_Click({ 
$mb = [System.Windows.Forms.MessageBox]
$mbIcon = [System.Windows.Forms.MessageBoxIcon]
$mbBtn = [System.Windows.Forms.MessageBoxButtons]


$result = $mb::Show("Are you sure you want reset user password?: " +$textInformacoesBox1.Text+"?", "AD Reset Password", $mbBtn::YesNo, $mbIcon::Exclamation)
If ($result -eq "Yes") {


    Try{

    $password = gerarNovaSenha
    $pass = ConvertTo-SecureString $password -AsPlainText -Force 
    Set-ADAccountPassword   -server $global:nomeDominio  -Identity $global:username -NewPassword $pass –Reset
    Set-ADUser -server $global:nomeDominio  -Identity $global:username -ChangePasswordAtLogon $true
    $TextBoxSenha.Text = $password
    $checkBox2.Checked = $true;


    #Chamando a função para atualizar o log
   logMensagem ($usuarioAtual+"::Performed that action::Password Reset::"+$global:username+"::"+$horaAtual)
 [System.Windows.MessageBox]::Show('Password reset performed','AD Reset Password',0,64)

   }Catch{
   
   [System.Windows.MessageBox]::Show("User not found",'AD Reset Password',0,64)
   
   }}
  
})

$buttonDesbloquear.Add_Click({ 

    Try{
    Unlock-ADAccount -server $global:nomeDominio -Identity $global:username
[System.Windows.MessageBox]::Show('User has been unlocked','Ferramenta Reset de senha - ATOS',0,64)
    logMensagem ($usuarioAtual+"::Performed that action::Unlock::"+$global:username+"::"+$horaAtual)
    
    }Catch{
    [System.Windows.MessageBox]::Show('User not found','AD Reset Password',0,64)
    }


})

$TextBoxUsuario.Add_KeyDown({
            if ($_.KeyCode -eq "Enter") {
                buscar
            }
    })

function verificaNulo {
   if($usuario -eq $null -or $usuario -eq ""){
         [System.Windows.MessageBox]::Show('User not found','AD Reset Password',0,64)
            }
            else{
            $textInformacoesBox1.Text =$usuario.name
            $TextBoxEmail.Text =$usuario.mail
            $TextBoxDepartamento.Text=$usuario.Department
            }
        }

Function buscar {
        

if($TextBoxUsuario.TextLength -eq 0){
         [System.Windows.MessageBox]::Show('Fill the user field','AD Reset Password',0,48)
        }
        else{

   #limparCampos
 $textInformacoesBox1.Text = ""
 $TextBoxEmail.Text=""
 $TextBoxDepartamento.Text=""

statusProximoLogin

$TextBoxSenha.Text = ""
$checkBox1.Checked = $false;
$checkBox2.Checked = $false;

Try
{
$usuario = get-aduser -server $global:nomeDominio  -identity $TextBoxUsuario.Text -properties * | select name,mail,Department, SamAccountName, enabled, lockedout, passwordexpired, pwdlastset 
$global:username = $usuario.SamAccountName
 
 #desbloquearBotao
 $buttonNovaSenha.Enabled = $true
 $buttonDesbloquear.Enabled = $true

   if($usuario.pwdlastset -eq 0){
     $checkBox2.Checked  = $true
     }
     
   if($usuario.lockedout -eq $true){
     $checkBox1.Checked  = $true
     }

verificaNulo

}
Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
{ 
 [System.Windows.MessageBox]::Show('User not found','AD Reset Password',0,64)
  #bloquearBotao
  $buttonNovaSenha.Enabled = $false
  $buttonDesbloquear.Enabled = $false
}

}

}
#Função que adiciona uma nova informação ao log de alteração.
Function logMensagem([String]$Message)
{
    Add-Content -Path "C:\Passwords\logs\log.txt" $Message
}


function gerarNovaSenha {
    param (
        [int]$length = 8,
        [bool]$includeLowercase = $true,
        [bool]$includeUppercase = $true,
        [bool]$includeNumbers = $true,
        [bool]$includeSpecialChars = $true
    )

    $lowercaseChars = "abcdefghijklmnopqrstuvwxyz"
    $uppercaseChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $numberChars = "0123456789"
    $specialChars = "!@*"

    $validChars = ""

    if ($includeLowercase) { $validChars += $lowercaseChars }
    if ($includeUppercase) { $validChars += $uppercaseChars }
    if ($includeNumbers) { $validChars += $numberChars }
    if ($includeSpecialChars) { $validChars += $specialChars }

    # Ensure at least one special character in the password
    $password = ""

    # Add a random special character
    $randomSpecialChar = Get-Random -Minimum 0 -Maximum $specialChars.Length
    $password += $specialChars[$randomSpecialChar]

    for ($i = 1; $i -lt $length; $i++) {
        $randomChar = Get-Random -Minimum 0 -Maximum $validChars.Length
        $password += $validChars[$randomChar]
    }

    # Shuffle the password to mix characters
    $passwordArray = $password.ToCharArray()
    $shuffledPassword = ($passwordArray | Get-Random -Count $passwordArray.Length) -join ""

    return $shuffledPassword
}


$panel.Controls.Add($TextBoxUsuario)
$panel.Controls.Add($Button)
$panel.Controls.Add($text1)
$panel.Controls.Add($textInformacoes1)
$panel.Controls.Add($textInformacoesBox1)
$panel.Controls.Add($textCheckBox1)
$panel.Controls.Add($checkBox1)
$panel.Controls.Add($textCheckBox2)
$panel.Controls.Add($checkBox2)
$panel.Controls.Add($textEmail)
$panel.Controls.Add($textBoxEmail)
$panel.Controls.Add($textDepartamento)
$panel.Controls.Add($textBoxDepartamento)

$panel.Controls.Add($textSenha)
$panel.Controls.Add($textBoxSenha)
$panel.Controls.Add($buttonNovaSenha)
$panel.Controls.Add($buttonDesbloquear)

$main_form.Controls.Add($TextBoxUsuario)
$main_form.Controls.Add($Button)
$main_form.Controls.Add($text1)

$main_form.Controls.Add($comboBoxDominio)
$main_form.Controls.Add($textDominio)
$main_form.Controls.Add($panel)

$main_form.ShowDialog()