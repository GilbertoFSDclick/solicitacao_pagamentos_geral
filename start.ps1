# Parametros
$arquivo_main = "main.py" # Nome do arquivo principal do bot
$caminho_projeto = $PSScriptRoot
$venv_python = Join-Path $PSScriptRoot "venv\Scripts\python.exe"

# Usa venv se existir; senao usa python do PATH
if (Test-Path $venv_python) {
    $caminho_python = $venv_python
    Write-Output "[PSScript] Usando venv: $caminho_python"
} else {
    $caminho_python = "python"
    Write-Output "[PSScript] Venv nao encontrado. Use setup.ps1 para criar. Usando python do PATH."
}
# $dclick_updater = "dclick-updater.exe"
# $caminho_updater = Join-Path $PSScriptRoot $dclick_updater

# Garante execucao no diretorio do script
Push-Location $PSScriptRoot

# ### UPDATER ###
# try {
#     & $caminho_updater
#     $exit_code = $LASTEXITCODE

#     if ($exit_code -ne 0) { throw }
# }
# catch {
#     Write-Output "[PSScript] Ocorreu um erro na execucao do updater: $_"
#     exit 1
# }

### BOT ###
try {
    $caminho_script = Join-Path $caminho_projeto $arquivo_main

    # Inicia o processo e espera terminar
    Write-Output "[PSScript] Iniciando automacao no caminho '$caminho_script'"
    
    & $caminho_python -u $caminho_script
    $exit_code = $LASTEXITCODE

    Write-Output "[PSScript] Execucao finalizada com codigo '$exit_code'"
    if ($exit_code -ne 0) { throw }
}
catch {
    Write-Output "[PSScript] Ocorreu um erro na execucao do bot: $_"
    exit 1
}
finally {
    Write-Output "[PSScript] Processo finalizado"
    [Console]::Out.Flush()
}