# Caminho para o arquivo Excel
$excelFilePath = "C:\Users\lgonzales\Desktop\shellscript\carlos\app\ad.xlsx"

# Nome da coluna que contém os nomes dos colaboradores
$columnName = "Nome do Colaborador"

# Carregar o módulo do Active Directory
Import-Module ActiveDirectory

# Ler o arquivo Excel e extrair os nomes dos colaboradores
$excelData = Import-Excel -Path $excelFilePath
$colaboradores = $excelData | Select-Object -ExpandProperty $columnName

# Iterar pelos nomes dos colaboradores e buscar no Active Directory
foreach ($nome in $colaboradores) {
    $user = Get-ADUser -Filter {DisplayName -eq $nome}

    if ($user) {
        Write-Host "Nome do Colaborador: $nome"
        Write-Host "Usuário encontrado no AD: $($user.SamAccountName)"
        Write-Host "-----------------------------"
    } else {
        Write-Host "Nome do Colaborador: $nome"
        Write-Host "Usuário não encontrado no AD"
        Write-Host "-----------------------------"
    }
}

