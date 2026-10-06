# Cópia para publicação: configure os exemplos abaixo antes de usar.
# ============================================================
# PROJETO: Limpeza de Histórico de Versões do SharePoint
# SCRIPT: 02 - Relatório detalhado de versões por pasta
# OBJETIVO: Ler arquivos e versões antigas e gerar CSV
# SEGURANÇA: Este script NÃO exclui nada
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------

$SiteUrl = "https://suaempresa.sharepoint.com/sites/Teste"

$Biblioteca = "Documentos"

# Se deixar vazio "", varre a biblioteca inteira
# Se preencher, varre somente a pasta informada
$Pasta = "/sites/Teste/Documentos/PastaTeste"

$PastaRelatorios = Join-Path $PSScriptRoot "Relatorios"

$ClientId = "" # Preencha localmente com o ID do seu aplicativo.

# Impede conexão enquanto os valores de exemplo não forem configurados.
if ($SiteUrl -eq "https://suaempresa.sharepoint.com/sites/Teste" -or
    [string]::IsNullOrWhiteSpace($ClientId)) {
    throw "Configure SiteUrl e ClientId localmente antes de executar este script."
}

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
# CONEXÃO
# ------------------------------------------------------------

function Connect-SharePoint {

    Write-Host ""
    Write-Host "Conectando ao SharePoint..."

    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        throw "ClientId ainda não foi configurado."
    }

    Connect-PnPOnline `
        -Url $SiteUrl `
        -Interactive `
        -ClientId $ClientId

    Write-Host "Conexão realizada."
}

# ------------------------------------------------------------
# INVENTÁRIO
# ------------------------------------------------------------

function Get-RelatorioVersoes {

    Write-Host ""
    Write-Host "========================================="
    Write-Host " RELATÓRIO DE VERSÕES DO SHAREPOINT"
    Write-Host "========================================="
    Write-Host ""

    Write-Host "Site: $SiteUrl"
    Write-Host "Biblioteca: $Biblioteca"

    if ([string]::IsNullOrWhiteSpace($Pasta)) {
        Write-Host "Escopo: biblioteca inteira"
    }
    else {
        Write-Host "Escopo da pasta:"
        Write-Host $Pasta
    }

    Write-Host ""

    $Resultado = @()

    # --------------------------------------------------------
    # BUSCA DOS ITENS
    # --------------------------------------------------------

    if ([string]::IsNullOrWhiteSpace($Pasta)) {

        $Itens = Get-PnPListItem `
            -List $Biblioteca `
            -PageSize 500 `
            -Fields "FileLeafRef","FileRef","FSObjType"
    }
    else {

        $Itens = Get-PnPListItem `
            -List $Biblioteca `
            -FolderServerRelativeUrl $Pasta `
            -PageSize 500 `
            -Fields "FileLeafRef","FileRef","FSObjType"
    }

    # --------------------------------------------------------
    # PROCESSAMENTO DOS ARQUIVOS
    # --------------------------------------------------------

    foreach ($Item in $Itens) {

        # FSObjType = 0 significa arquivo
        if ($Item.FieldValues["FSObjType"] -ne 0) {
            continue
        }

        $NomeArquivo = $Item.FieldValues["FileLeafRef"]
        $CaminhoArquivo = $Item.FieldValues["FileRef"]

        Write-Host "Analisando: $NomeArquivo"

        try {

            $Versoes = @(
                Get-PnPFileVersion `
                    -Url $CaminhoArquivo
            )

            $QuantidadeVersoes = $Versoes.Count

            if ($QuantidadeVersoes -eq 0) {

                $Resultado += [PSCustomObject]@{
                    Arquivo    = $NomeArquivo
                    Caminho    = $CaminhoArquivo
                    Versao     = ""
                    IDVersao   = ""
                    DataVersao = ""
                    Autor      = ""
                    TamanhoMB  = 0
                    Acao       = "Nenhuma ação"
                }

            }
            else {

                foreach ($Versao in $Versoes) {

                    $TamanhoMB = [math]::Round(
                        ($Versao.Size / 1MB),
                        2
                    )

                    $AutorVersao = ""

                    if ($null -ne $Versao.CreatedBy) {
                        $AutorVersao = $Versao.CreatedBy.LoginName
                    }

                    $Resultado += [PSCustomObject]@{
                        Arquivo    = $NomeArquivo
                        Caminho    = $CaminhoArquivo
                        Versao     = $Versao.VersionLabel
                        IDVersao   = $Versao.Id
                        DataVersao = $Versao.Created
                        Autor      = $AutorVersao
                        TamanhoMB  = $TamanhoMB
                        Acao       = "Candidato a limpeza"
                    }
                }
            }
        }
        catch {

            $Resultado += [PSCustomObject]@{
                Arquivo    = $NomeArquivo
                Caminho    = $CaminhoArquivo
                Versao     = ""
                IDVersao   = ""
                DataVersao = ""
                Autor      = ""
                TamanhoMB  = ""
                Acao       = "ERRO: $($_.Exception.Message)"
            }
        }
    }

    return $Resultado
}

# ------------------------------------------------------------
# EXECUÇÃO
# ------------------------------------------------------------

try {

    Connect-SharePoint

    $Relatorio = @(Get-RelatorioVersoes)

    if ($Relatorio.Count -gt 0) {

        $Relatorio |
            Export-Csv `
                -Path $ArquivoRelatorio `
                -NoTypeInformation `
                -Delimiter ";" `
                -Encoding utf8BOM

        Write-Host ""
        Write-Host "Relatório criado:"
        Write-Host $ArquivoRelatorio
    }
    else {

        Write-Host ""
        Write-Host "Nenhum arquivo encontrado."
    }
}
catch {

    Write-Host ""
    Write-Host "ERRO:"
    Write-Host $_.Exception.Message
}
finally {

    Disconnect-PnPOnline -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "Fim da execução."
}
