<#
.DESCRIPTION
  Script to import list of user using CSV File
    Need to fill in $UPN  on line 104 
.NOTES
    NAME: AD Import User from CSV 
    AUTHOR: Matheus Andrade
    LASTEDIT: 09/09/2023 

#>

[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName 'System.Web'
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles();
#Criando o formulário 
$form = New-Object Windows.Forms.Form
$form.Text = "AD Import User from CSV"
$form.Size = New-Object Drawing.Size(600, 200)
$form.StartPosition = "CenterScreen"
$Font = New-Object System.Drawing.Font("Verdana",12)
$form.Font = $Font
$form.BackColor = "white"
$form.MaximizeBox = $false
$form.FormBorderStyle = 'Fixed3D'

#Variavel Global que armazena o objeto com dados dos usuários
 $global:ADUsers

# Criando o label com caminho oda pasta
$label = New-Object Windows.Forms.Label
$label.Text = "Patch:"
$label.Width = 90
$label.Location = New-Object Drawing.Point(20,30)
$form.Controls.Add($label)

#Criando o textbox do caminho
$textbox = New-Object Windows.Forms.TextBox
$textbox.Location = New-Object Drawing.Point(110,30)
$textbox.Width = 250
$form.Controls.Add($textbox)

#Criando  o botão Abrir 
$button = New-Object Windows.Forms.Button
$button.Text = "Open"
$button.Location = New-Object Drawing.Point(380, 30)
$button.Size= "180,30"

#Criando  o botão Salvar 
$buttonSave = New-Object Windows.Forms.Button
$buttonSave.Text = "Import Users"
$buttonSave.Location = New-Object Drawing.Point(380, 70)
$buttonSave.Size= "180,30"
$buttonSave.Enabled = $false

$menubar = New-Object System.Windows.Forms.MenuStrip


#Evento do botão de carrega o arquivo CSV 
$button.Add_Click({

    $openFileDialog = New-Object Windows.Forms.OpenFileDialog
    $openFileDialog.Filter =  "CSV files (*.csv)|*.csv";
    $result = $openFileDialog.ShowDialog()
    
    if ($result -eq [Windows.Forms.DialogResult]::OK) {
        $selectedFile = $openFileDialog.FileName
        $textbox.Text = $selectedFile
    
     $global:ADUsers   = Import-Csv -Path $selectedFile  -Delimiter ";"

     $buttonSave.Enabled = $true
        
    }
})

#Evento do Botão que persiste os usuários
$buttonSave.Add_Click({
        cadastrar
    })


#Criando o menubar 
$helpMenu = $menubar.Items.Add("About")

$helpMenu.Add_Click({

#Mensagem da opção Sobre
$Msg = @'
Developed by Matheus Andrade
github.com/matheusandradeX
'@

   [System.Windows.MessageBox]::Show($Msg,'AD Import User from CSV',0,64)
})


Function cadastrar{
# Define UPN necessário preencher
$UPN = "ambiente.homologacao"

# Loop through each row containing user details in the CSV file
foreach ($User in  $global:ADUsers) {

    #Read user data from each field in each row and assign the data to a variable as below
    $username = $User.Username
    $password = $User.password
    $firstname = $User.firstname
    $lastname = $User.lastname
    $initials = $User.initials
    $OU = $User.ou #This field refers to the OU the user account is to be created in
    $email = $User.email
    $streetaddress = $User.streetaddress
    $city = $User.city
    $zipcode = $User.zipcode
    $state = $User.state
    $country = $User.country
    $telephone = $User.telephone
    $jobtitle = $User.jobtitle
    $company = $User.company
    $department = $User.department

    #Verificando se usuário possui conta no AD
    if (Get-ADUser -F { SamAccountName -eq $username }) {
         [System.Windows.MessageBox]::Show("The $username account already exists",'AD Import User from CSV',0,48)
        Write-Warning "The user account $username already exists in Active Directory"
    }
    else {

        #Caso não exista o usuário será criado no Active Diretory 
        New-ADUser `
            -SamAccountName $username `
            -UserPrincipalName "$username@$UPN" `
            -Name "$firstname $lastname" `
            -GivenName $firstname `
            -Surname $lastname `
            -Initials $initials `
            -Enabled $True `
            -DisplayName "$lastname, $firstname" `
            -Path $OU `
            -City $city `
            -PostalCode $zipcode `
            -Country $country `
            -Company $company `
            -State $state `
            -StreetAddress $streetaddress `
            -OfficePhone $telephone `
            -EmailAddress $email `
            -Title $jobtitle `
            -Department $department `
            -AccountPassword (ConvertTo-secureString $password -AsPlainText -Force) -ChangePasswordAtLogon $True
       [System.Windows.MessageBox]::Show("The $username account has been created",'AD Import User from CSV',0,64)
       
    }
}

}


$form.Controls.Add($menubar)
$form.Controls.Add($button)
$form.Controls.Add($buttonSave)
$form.ShowDialog()
