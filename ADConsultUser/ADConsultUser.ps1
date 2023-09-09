<#
.SYNOPSIS
  Active Diretory only read for service desk team 
.DESCRIPTION
  A simple tool to consult user attributes using like Active Diretory
.NOTES
  Version:        .9
  Author:         Matheus Andrade
  Creation Date:  2022
#>

#Código para ocultar o prompt de comando do powershell após a incialização
#$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);
#add-type -name win -member $t -namespace native
#[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
#Inicialização do formulário
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles();
#Inicialização do formulário
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
#Declaração de elementos da GUI 

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Consult User AD'
$main_form.Width = 585
$main_form.Height = 460
$main_form.BackColor = "white"
$main_form.MaximizeBox = $false
$main_form.FormBorderStyle = 'Fixed3D'
$main_form.StartPosition = "CenterScreen"
$Font = New-Object System.Drawing.Font("Verdana",9)
$main_form.Font = $Font
$text1 = New-Object System.Windows.Forms.label
$text1.Location = New-Object System.Drawing.Size(7,2)
$text1.Size = New-Object System.Drawing.Size(290,15)
$text1.ForeColor = "black"
$text1.Text = "Type the employee's name : "
$text2 = New-Object System.Windows.Forms.label
$text2.Location = New-Object System.Drawing.Size(7,395)
$text2.Size = New-Object System.Drawing.Size(270,15)
$text2.ForeColor = "black"
$text2.Text = " "
$TextBoxPatch = New-Object System.Windows.Forms.TextBox
$TextBoxPatch.Location = '10,26'
$TextBoxPatch.Size = '370,50'
#Variavel global do usuário de rede para pesquisa
$global:textField 
#Função para chamar a pesquisa geral ao prescionar Enter no campo Nome
$TextBoxPatch.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        pesquisaGeral
    }
})
#Declaração de botões na tela inical
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(388,25)
$Button.Size = New-Object System.Drawing.Size(170,25)
$Button.Text = "Search"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(388,55)
$Button2.Size = New-Object System.Drawing.Size(170,25)
$Button2.Text = "Serach by username"
#Delcaração do Lista de elementos 
$listBox = New-Object System.Windows.Forms.ListView
$listBox.Location = New-Object System.Drawing.Point(15,90)
$listBox.Size = New-Object System.Drawing.Size(540,20)
$listBox.Height = 300
$listBox.Columns.Add("User Name")
$listBox.Columns.Add("Name")
$listBox.Columns.Add("Status")
$listBox.Columns[0].Width = 90
$listBox.Columns[1].Width = 360
$listBox.View = [System.Windows.Forms.View]::Details
$listBox.MultiSelect = $false
$listBox.FullRowSelect = $true
#Função que após prescionar o botão pesquisar chamar a função pesquisa Geral
$Button.Add_Click({ 
    pesquisaGeral
})
#Função do botão buscar por usuário que incializa a tela de pesquisa por usuário de rede, e limpa se tiver essa tela estiver preenchida 
$Button2.Add_Click({
$listBox.SelectedItems.Clear()
buscarUsuario

})
#Instanciando formuário da tela inicial
$listBox.Add_ItemActivate({selecionarLinha})
#Função para identificar qual linha foi precionanda com 2 clikes 
Function selecionarLinha(){
    if($listBox.SelectedItems.Count -eq 1){
    $usuarioRedeLinhaSelecionada = $listBox.SelectedItems[0].Text
   buscarUsuario
    }
}
#Função que pesquisa usuário por meio do nome
function pesquisaGeral{
#limpando o list view caso anteriomente foi preenchido
$listBox.Items.Clear();
#Pegando o valor do nome do campo de input
$nome =  $TextBoxPatch.Text
#Adicionando os caracteres especiais entre o nome para poder fazer a buscar sem precisar do nome exato 

$nomeAsterisco=$nome.Replace(" ","*")

$var = -join("*", $nomeAsterisco, "*");
#incializado a variavel contador
$contador = 0
#Verificação se de pesquisa não está vazia
if($TextBoxPatch.TextLength -eq 0){
     [System.Windows.MessageBox]::Show('Fill the name field!','AD User Consult',0,48)
    }else{
    #Consulta ao Active Diretory para buscar , usuário de rede, nome e status 
$dados = ForEach-Object {Get-Aduser -Filter {name -like $var } -properties name,samaccountname,enabled   | select name,samaccountname,enabled   }

#Laço de repetição para percorrer o resultado da busca
foreach ($dado in $dados) {

#Conventendo o valor da variavel enabled para string pois list view não suporta valores boolean
if($dado.enabled -eq $true){
   $status= "Enabled"
}else {
    $status= "Disabled"
}

#Adcionando cada valor em sua respecitiva coluna
$listBox.Items.Add($dado.samaccountname)
$listBox.Items[$contador].SubItems.Add($dado.name);
$listBox.Items[$contador].SubItems.Add($status);
#Varivael para percorrer 2 coluna em diante 
$contador +=1
} 

#Contador de resultados encontrados 
$text2.Text = 'Results found: '+$contador
}}
Function buscarUsuario{
#Inicialização do formulário
[System.Windows.Forms.Application]::EnableVisualStyles();
#Declaração de elementos da GUI 
Add-Type -Assembly 'System.Windows.Forms'
$form = New-Object Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(370,770) 
$form.StartPosition = "CenterScreen"
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Size(10,10)
$TabControl.Size = New-Object System.Drawing.Size(330,710)
$form.FormBorderStyle = 'Fixed3D'
$form.Text ='AD User Consult'
$tabPage1 = New-Object System.Windows.Forms.TabPage
$tabPage1.Name = "Tab1"
$tabPage1.Text = "Employee Info"
$tabPage2 = New-Object System.Windows.Forms.TabPage
$tabPage2.Name = "Tab2"
$tabPage2.Text = "Group Membership"
$tabPage1.BackColor = "white"
$form.MaximizeBox = $false
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 20,10
$label.Size = New-Object Drawing.Point 200,15
$label.text = "Type the user's name:"
$global:textField = New-Object Windows.Forms.TextBox
$global:textField.Location = New-Object Drawing.Point 20,30
$global:textField.Size = New-Object Drawing.Point 250,15
$global:textField.Text = $listBox.SelectedItems[0].Text
$Font = New-Object System.Drawing.Font("Verdana",9)
$form.Font = $Font
#Variavel Global que armazena os dados do usuário
$global:user
#Verificação se encontrou o usuário
  function verificaNulo {
   if($global:user -eq $null -or $global:user -eq ""){
         [System.Windows.MessageBox]::Show('User did not find','AD User Consult')
            }
        }
        $textField.Add_KeyDown({
            if ($_.KeyCode -eq "Enter") {
                buscarDados
            }
    })
#Delcaração dos campos do formulário
$labelMatricula = New-Object Windows.Forms.Label
$labelMatricula.Location = New-Object Drawing.Point 20,60
$labelMatricula.Size = New-Object Drawing.Point 130,20
$labelMatricula.text = "EMPLOYEE ID"
$textFieldMatricula = New-Object Windows.Forms.TextBox
$textFieldMatricula.Location = New-Object Drawing.Point 20,80
$textFieldMatricula.Size = New-Object Drawing.Point 250,20
$textFieldMatricula.ReadOnly = "true"
$labelNome = New-Object Windows.Forms.Label
$labelNome.Location = New-Object Drawing.Point 20,110
$labelNome.Size = New-Object Drawing.Point 130,20
$labelNome.text = "FULL NAME "
$textFieldNome = New-Object Windows.Forms.TextBox 
$textFieldNome.Location = New-Object Drawing.Point 20,130
$textFieldNome.Size = New-Object Drawing.Point 250,20
$textFieldNome.ReadOnly = "true"
$labelEmail = New-Object Windows.Forms.Label
$labelEmail.Location = New-Object Drawing.Point 20,160
$labelEmail.Size = New-Object Drawing.Point 250,20
$labelEmail.text = "E-MAIL"
$textFieldEmail = New-Object Windows.Forms.TextBox
$textFieldEmail.Location = New-Object Drawing.Point 20,180
$textFieldEmail.Size = New-Object Drawing.Point 250,20
$textFieldEmail.ReadOnly = "true"
$labelCargo = New-Object Windows.Forms.Label
$labelCargo.Location = New-Object Drawing.Point 20,210
$labelCargo.Size = New-Object Drawing.Point 250,20
$labelCargo.text = "JOB TITLE"
$textFieldCargo = New-Object Windows.Forms.TextBox
$textFieldCargo.Location = New-Object Drawing.Point 20,230
$textFieldCargo.Size = New-Object Drawing.Point 250,20
$textFieldCargo.ReadOnly = "true"
$labelDepartamento = New-Object Windows.Forms.Label
$labelDepartamento.Location = New-Object Drawing.Point 20,260
$labelDepartamento.Size = New-Object Drawing.Point 250,20
$labelDepartamento.text = "DEPARTMENT"
$textFieldDepartamento = New-Object Windows.Forms.TextBox
$textFieldDepartamento.Location = New-Object Drawing.Point 20,280
$textFieldDepartamento.Size = New-Object Drawing.Point 250,20
$textFieldDepartamento.ReadOnly = "true"
$labelGestor = New-Object Windows.Forms.Label
$labelGestor.Location = New-Object Drawing.Point 20,310
$labelGestor.Size = New-Object Drawing.Point 250,20
$labelGestor.text = "MANAGER"
$textFieldGestor = New-Object Windows.Forms.TextBox
$textFieldGestor.Location = New-Object Drawing.Point 20,330
$textFieldGestor.Size = New-Object Drawing.Point 250,20
$textFieldGestor.ReadOnly = "true"
$labelCPF = New-Object Windows.Forms.Label
$labelCPF.Location = New-Object Drawing.Point 20,360
$labelCPF.Size = New-Object Drawing.Point 250,20
$labelCPF.text = "DESCRIPTION"
$textFieldCPF = New-Object Windows.Forms.TextBox
$textFieldCPF.Location = New-Object Drawing.Point 20,380
$textFieldCPF.Size = New-Object Drawing.Point 250,20
$textFieldCPF.ReadOnly = "true"
$labelCC = New-Object Windows.Forms.Label
$labelCC.Location = New-Object Drawing.Point 20,410
$labelCC.Size = New-Object Drawing.Point 130,20
$labelCC.text = "COST CENTER"
$textFieldCC = New-Object Windows.Forms.TextBox
$textFieldCC.Location = New-Object Drawing.Point 20,430
$textFieldCC.Size = New-Object Drawing.Point 250,20
$textFieldCC.ReadOnly = "true"
$labelRamal = New-Object Windows.Forms.Label
$labelRamal.Location = New-Object Drawing.Point 20,460
$labelRamal.Size = New-Object Drawing.Point 250,20
$labelRamal.text = "TELEPHONE"
$textFieldRamal = New-Object Windows.Forms.TextBox
$textFieldRamal.Location = New-Object Drawing.Point 20,480
$textFieldRamal.Size = New-Object Drawing.Point 250,20
$textFieldRamal.ReadOnly = "true"
$labelLocalidade = New-Object Windows.Forms.Label
$labelLocalidade.Location = New-Object Drawing.Point 20,510
$labelLocalidade.Size = New-Object Drawing.Point 130,20
$labelLocalidade.text = "LOCATION"
$textFieldLocalidade = New-Object Windows.Forms.TextBox
$textFieldLocalidade.Location = New-Object Drawing.Point 20,530
$textFieldLocalidade.Size = New-Object Drawing.Point 250,20
$textFieldLocalidade.ReadOnly = "true"
$labelEmpresa = New-Object Windows.Forms.Label
$labelEmpresa.Location = New-Object Drawing.Point 20,560
$labelEmpresa.Size = New-Object Drawing.Point 130,20
$labelEmpresa.text = "COMPANY"
$textFieldEmpresa = New-Object Windows.Forms.TextBox
$textFieldEmpresa.Location = New-Object Drawing.Point 20,580
$textFieldEmpresa.Size = New-Object Drawing.Point 250,20
$textFieldEmpresa.ReadOnly = "true"
$button = New-Object Windows.Forms.Button
$button.text = "SEARCH"
$button.Location = New-Object Drawing.Point 20,650
$buttonCopiar = New-Object Windows.Forms.Button
$buttonCopiar.text = "COPY"
$buttonCopiar.Location = New-Object Drawing.Point 195,650
$textCheckBox = New-Object System.Windows.Forms.label
$textCheckBox.Location = New-Object System.Drawing.Size(35,617)
$textCheckBox.Size = New-Object System.Drawing.Size(290,15)
$textCheckBox.Text = "Must change password at next logon"
$checkbox1 = new-object System.Windows.Forms.checkbox
$checkbox1.Location = new-object System.Drawing.Size(20,610)
$checkbox1.Size = new-object System.Drawing.Size(30,30)
$checkBox1.Enabled = $false;


    #TabPage2 Grupo de acessos
    $tabPage2.BackColor = "white"
    # user label
    $grupoAcessosLabel = New-Object System.Windows.Forms.label
    $grupoAcessosLabel.Location = New-Object System.Drawing.Size(10,13)
    $grupoAcessosLabel.Size = New-Object System.Drawing.Size(170,25)
    $grupoAcessosLabel.Text = "Access Groups:"
    #Delcaração do Lista de elementos 
    $listBoxGrupos = New-Object System.Windows.Forms.ListView
    #posicao do campo de lista de grupos
    $listBoxGrupos.Location = New-Object System.Drawing.Point(10,40) 
    $listBoxGrupos.Size = New-Object System.Drawing.Size(300,600)
    $listBoxGrupos.Height = 630
    $listBoxGrupos.Columns.Add("Name")
    $listBoxGrupos.Columns[0].Width = 275
    $listBoxGrupos.View = [System.Windows.Forms.View]::Details
    $buttonGrupos = New-Object Windows.Forms.Button
    $buttonGrupos.text = "Search"
    $buttonGrupos.Size = New-Object Drawing.Point 90,22
    $buttonGrupos.Location = New-Object Drawing.Point 220,10
    # Adicionado elementos no form tabpage2.
    $tabPage2.Controls.AddRange($listBoxGrupos)
    $tabPage2.Controls.AddRange($grupoAcessosLabel)
    $tabPage2.Controls.AddRange($buttonGrupos)
   
   $buttonGrupos.add_click({
    buscarGruposAcessos
       })

    function buscarGruposAcessos {
    
        $listBoxGrupos.Items.Clear();
        if($textField.Text){
      #Consulta ao Active Diretory para buscar , usuário de rede, nome e status 
    $dados = ForEach-Object {Get-ADPrincipalGroupMembership  $textField.Text | Select-Object -ExpandProperty name  }
    for($count=0;$count -lt $dados.Length;$count++){
    $listBoxGrupos.Items.Add($dados[$count])
    }
    }
    
    }


    #Verificação se o campo usuário de rede foi prenchido na tela anterior por meio linha selecicionada, caso sim é pesquisado as informações do colaborador e é preenchido o formulário
    if($textField.Text){
    $checkBox1.Checked  = $false
    $global:user = Get-Aduser -Filter {samaccountname -eq $textField.Text } -properties GivenName,SurName, mail , Title, Manager, Department,Description,Departmentnumber,EmployeeID,Company,telephoneNumber,l,pwdlastset | Select-Object  GivenName,SurName, mail , Title, Manager, Department,Description,@{N='Departmentnumber';E={$_.Departmentnumber[0]}},EmployeeID,Company,telephoneNumber,l,pwdlastset
    verificaNulo

     if($user.Manager -eq $null -or $user.Manager -eq ""){
     $getorSomenteNome = ""
    
     }else {
     
      $gestorSplit = $user.Manager.Split(',')[0]
      $getorSomenteNome = $gestorSplit.Split('=')[1]
     }
    
    if($user.pwdlastset -eq 0){
     $checkBox1.Checked  = $true
     }
    
    for($count=0;$count -lt $dados.Length;$count++){
    $listBox.Items.Add($dados[$count])
    }
    #limpar dados lista de grupos
    $listBoxGrupos.Items.Clear();
    $nomeCompleto = -join($user.GivenName," ", $user.SurName);
    $textFieldNome.Text = $nomeCompleto
    $textFieldEmail.Text = $user.mail
    $textFieldCargo.Text = $user.Title
    $textFieldDepartamento.Text = $user.Department
    $textFieldGestor.Text = $getorSomenteNome
    $textFieldCPF.Text = $user.Description
    $textFieldCC.Text = $user.Departmentnumber
    $textFieldMatricula.Text = $user.EmployeeID
    $textFieldEmpresa.Text = $user.Company
    $textFieldRamal.Text=$user.telephoneNumber
    $textFieldLocalidade.Text = $user.l
    }
    #Evendo do botão buscar na tela buscar por usuário de rede
    $button.add_click({
    buscarDados
       })
    #Função que pesquisa as informações do colaborador e é preenchido o formulário
    function buscarDados{
    #Veirifcação se o campo usuário foi preenchido antes de realizar a busca
        if($textField.TextLength -eq 0){
         [System.Windows.MessageBox]::Show('Fill the user name field!','AD User Consult')
        }else{
        $checkBox1.Checked  = $false
        $global:user = Get-Aduser -Filter {samaccountname -eq $textField.Text } -properties GivenName,SurName, mail , Title, Manager, Department,Description,Departmentnumber,EmployeeID,Company,telephoneNumber,l,pwdlastset | Select-Object  GivenName,SurName, mail , Title, Manager, Department,Description,@{N='Departmentnumber';E={$_.Departmentnumber[0]}},EmployeeID,Company,telephoneNumber,l,pwdlastset
        verificaNulo
         if($user.Manager -eq $null -or $user.Manager -eq ""){
        $getorSomenteNome = ""
     }else {
     
      $gestorSplit = $user.Manager.Split(',')[0]
      $getorSomenteNome = $gestorSplit.Split('=')[1]
     }

     if($user.pwdlastset -eq 0){
     $checkBox1.Checked  = $true
     }
         #limpar dados lista de grupos
        $listBoxGrupos.Items.Clear();
        $nomeCompleto = -join($user.GivenName," ", $user.SurName);
        $textFieldNome.Text = $nomeCompleto
        $textFieldEmail.Text = $user.mail
        $textFieldCargo.Text = $user.Title
        $textFieldDepartamento.Text = $user.Department
        $textFieldGestor.Text = $getorSomenteNome
        $textFieldCPF.Text = $user.Description
        $textFieldCC.Text = $user.Departmentnumber
        $textFieldMatricula.Text = $user.EmployeeID
        $textFieldEmpresa.Text = $user.Company
        $textFieldRamal.Text=$user.telephoneNumber
        $textFieldLocalidade.Text = $user.l
            }
        }
    #Função para copiar elementos para área de transferência
    function copiarDados{
        $textFieldCopy = "User name: "+$textField.Text
        $nomeCompletoCopy = "Full Name: "+$textFieldNome.Text
        $textFieldEmailCopy = "E-mail: "+ $textFieldEmail.Text
        $textFieldMatriculaCopy ="Employee ID: "+$textFieldMatricula.Text 
        $textFieldCargoCopy = "Job Title: "+$textFieldCargo.Text
        $textFieldDepartamentoCopy = "Department: "+$textFieldDepartamento.Text
        $textFieldCPFCopy = "Description: "+$textFieldCPF.Text
        $textFieldCCCopy = "Cost Center: "+$textFieldCC.Text
        $textFieldRamalCopy = "Telephone: "+$textFieldRamal.Text
        $textFieldLocalidadeCopy = "Location: "+$textFieldLocalidade.Text
        $textFieldEmpresaCopy = "Company: "+$textFieldEmpresa.Text
        Set-Clipboard $nomeCompletocopy,$textFieldMatriculaCopy,$textFieldCopy,$textFieldEmailCopy,$textFieldCargoCopy,$textFieldDepartamentoCopy,$textFieldCPFCopy,$textFieldCCCopy,$textFieldRamalCopy,$textFieldLocalidadeCopy,$textFieldEmpresaCopy 
        [System.Windows.MessageBox]::Show('information copy to Clipboard','AD User Consult',0,64)
    }
    $buttonCopiar.add_click({
      copiarDados
       })
    # Adicionado elementos no form tabpage1.
    $tabPage1.Controls.AddRange($button)
    $tabPage1.Controls.AddRange($buttonCopiar)
    $tabPage1.Controls.AddRange($label)
    $tabPage1.Controls.AddRange($textField)
    $tabPage1.Controls.AddRange($labelMatricula)
    $tabPage1.Controls.AddRange($textfieldMatricula)
    $tabPage1.Controls.AddRange($labelNome)
    $tabPage1.Controls.AddRange($textfieldNome)
    $tabPage1.Controls.AddRange($labelEmail)
    $tabPage1.Controls.AddRange($textfieldEmail)
    $tabPage1.Controls.AddRange($labelCargo)
    $tabPage1.Controls.AddRange($textfieldCargo)
    $tabPage1.Controls.AddRange($labelDepartamento)
    $tabPage1.Controls.AddRange($textfieldDepartamento)
    $tabPage1.Controls.AddRange($labelGestor)
    $tabPage1.Controls.AddRange($textfieldGestor)
    $tabPage1.Controls.AddRange($labelCPF)
    $tabPage1.Controls.AddRange($textfieldCPF)
    $tabPage1.Controls.AddRange($labelCC)
    $tabPage1.Controls.AddRange($textfieldCC)
    $tabPage1.Controls.AddRange($labelRamal)
    $tabPage1.Controls.AddRange($textfieldRamal)
    $tabPage1.Controls.AddRange($labelLocalidade)
    $tabPage1.Controls.AddRange($textfieldLocalidade)
    $tabPage1.Controls.AddRange($labelEmpresa)
    $tabPage1.Controls.AddRange($textfieldEmpresa)
    $tabPage1.Controls.AddRange($textCheckBox)
    $tabPage1.Controls.AddRange($checkbox1)
    $TabControl.TabPages.Add($tabPage1)
    $TabControl.TabPages.Add($tabPage2) 
    $form.Controls.Add($TabControl)
    $form.Topmost = $True
    $form.Add_Shown({$form.Activate()})
    [void] $form.ShowDialog()
}
$main_form.Controls.Add($text1)
$main_form.Controls.Add($text2)
$main_form.Controls.Add($TextBoxPatch)
$main_form.Controls.Add($Button)
$main_form.Controls.Add($Button2)
$main_form.Controls.Add($listBox)
$main_form.ShowDialog()