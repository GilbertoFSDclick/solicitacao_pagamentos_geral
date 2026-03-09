# Setup - Cria venv e instala dependencias
# Execute: .\setup.ps1

$caminho_projeto = $PSScriptRoot
$venv_dir = Join-Path $caminho_projeto "venv"

Push-Location $caminho_projeto

if (Test-Path $venv_dir) {
    Write-Output "[Setup] Venv ja existe em $venv_dir"
} else {
    Write-Output "[Setup] Criando venv..."
    python -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Error "[Setup] Erro ao criar venv. Verifique se Python esta instalado e no PATH."
        exit 1
    }
    Write-Output "[Setup] Venv criado."
}

Write-Output "[Setup] Instalando dependencias..."
& "$venv_dir\Scripts\pip.exe" install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error "[Setup] Erro ao instalar dependencias."
    exit 1
}

Write-Output "[Setup] Concluido. Execute start.ps1 para rodar o bot."
Pop-Location
