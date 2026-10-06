# Cópia para publicação: configure os exemplos abaixo antes de usar.
# ============================================================
# PROJETO: Limpeza de Histórico de Versões do SharePoint
# SCRIPT: 01 - Inventário de versões
# OBJETIVO: Ler arquivos e identificar versões antigas
# SEGURANÇA: Este script NÃO exclui nada
# ============================================================

Set-StrictMode -Version Latest

# ------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------

$SiteUrl = "https://suaempresa.sharepoint.com/sites/Teste"

$Biblioteca = "Documentos"

$Pasta = ""

$PastaRelatorios = Join-Path $PSScriptRoot "Relatorios"

# ------------------------------------------------------------
# PREPARAÇÃO
# ------------------------------------------------------------

if (-not (Test-Path $PastaRelatorios)) {

    New-Item `
        -ItemType Directory `
        -Path $PastaRelatorios `
        -Force | Out-Null
}

$DataExecucao = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$ArquivoRelatorio =
    "$PastaRelatorios\Relatorio_Versoes_$DataExecucao.csv"

# ------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ------------------------------------------------------------

function Get-RelatorioVersoes {

    Write-Host ""
    Write-Host "========================================="
    Write-Host " INVENTÁRIO DE VERSÕES DO SHAREPOINT"
    Write-Host "========================================="
    Write-Host ""

    Write-Host "Site: $SiteUrl"
    Write-Host "Biblioteca: $Biblioteca"
    Write-Host ""

    # Futuramente receberá os arquivos do SharePoint
    $Resultado = @()

    Write-Host "Estrutura do relatório preparada."
    Write-Host "Nenhum arquivo foi alterado."
    Write-Host ""

    return $Resultado
}

# ------------------------------------------------------------
# EXECUÇÃO
# ------------------------------------------------------------

$Relatorio = @(Get-RelatorioVersoes)

if ($Relatorio.Count -gt 0) {

    $Relatorio |
        Export-Csv `
            -Path $ArquivoRelatorio `
            -NoTypeInformation `
            -Encoding UTF8

    Write-Host "Relatório criado:"
    Write-Host $ArquivoRelatorio
}
else {

    Write-Host "Nenhum dado encontrado."
    Write-Host "Isso é esperado enquanto não conectarmos ao SharePoint."
}

Write-Host ""
Write-Host "Fim da execução."
