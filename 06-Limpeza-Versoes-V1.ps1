# Cópia para publicação: configure os exemplos abaixo antes de usar.
# ============================================================
# PROJETO: Automação de Limpeza de Versões do SharePoint
# SCRIPT: 06 - Limpeza de Versões V1.0
#
# MODOS:
# SIMULACAO = identifica e registra, mas NÃO exclui
# EXECUCAO  = envia versões históricas para a Lixeira
#
# SEGURANÇA:
# - Site definido
# - Biblioteca definida
# - Pasta obrigatória para EXECUCAO
# - Confirmação manual
# - Revalidação da versão antes da exclusão
# - Log detalhado
# - Tratamento individual de erros
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------

$SiteUrl = "https://suaempresa.sharepoint.com/sites/Teste"

$Biblioteca = "Documentos"

$Pasta = "/sites/Teste/Documentos/PastaTeste"

$ClientId = "" # Preencha localmente com o ID do seu aplicativo.

$PastaRelatorios = Join-Path $PSScriptRoot "Relatorios"

# ------------------------------------------------------------
# MODO
# ------------------------------------------------------------
# PRIMEIRO TESTE:
$ModoExecucao = "SIMULACAO"

# Para executar de verdade depois:
# $ModoExecucao = "EXECUCAO"

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

$ArquivoLog =
    "$PastaRelatorios\Log_Limpeza_V1_$DataExecucao.csv"

$Log = @()
$Candidatos = @()

$TotalArquivos = 0
$TotalVersoes = 0
$TotalSucesso = 0
$TotalErro = 0
$TotalIgnorado = 0
$EspacoTotalMB = 0

# ------------------------------------------------------------
# FUNÇÃO DE LOG
# ------------------------------------------------------------

function Add-Log {

    param (
        $Arquivo,
        $Caminho,
        $Versao,
        $IDVersao,
        $DataVersao,
        $TamanhoMB,
        $Resultado,
        $Mensagem
    )

    $script:Log += [PSCustomObject]@{

        DataExecucao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        Modo         = $ModoExecucao
        Arquivo      = $Arquivo
        Caminho      = $Caminho
        Versao       = $Versao
        IDVersao     = $IDVersao
        DataVersao   = $DataVersao
        TamanhoMB    = $TamanhoMB
        Resultado    = $Resultado
        Mensagem     = $Mensagem
    }
}

# ------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# ------------------------------------------------------------

try {

    Write-Host ""
    Write-Host "=================================================="
    Write-Host " SHAREPOINT - LIMPEZA DE VERSÕES V1.0"
    Write-Host "=================================================="
    Write-Host ""

    # --------------------------------------------------------
    # VALIDAR MODO
    # --------------------------------------------------------

    if (
        $ModoExecucao -ne "SIMULACAO" -and
        $ModoExecucao -ne "EXECUCAO"
    ) {

        throw "Modo inválido. Use SIMULACAO ou EXECUCAO."
    }

    # --------------------------------------------------------
    # TRAVA DE SEGURANÇA
    # --------------------------------------------------------

    if (
        $ModoExecucao -eq "EXECUCAO" -and
        [string]::IsNullOrWhiteSpace($Pasta)
    ) {

        throw "SEGURANÇA: EXECUCAO não permitida com a pasta vazia."
    }

    Write-Host "Modo:"
    Write-Host $ModoExecucao
    Write-Host ""

    Write-Host "Site:"
    Write-Host $SiteUrl
    Write-Host ""

    Write-Host "Biblioteca:"
    Write-Host $Biblioteca
    Write-Host ""

    Write-Host "Pasta:"
    Write-Host $Pasta
    Write-Host ""

    # --------------------------------------------------------
    # CONEXÃO
    # --------------------------------------------------------

    Write-Host "Conectando ao SharePoint..."

    Connect-PnPOnline `
        -Url $SiteUrl `
        -Interactive `
        -ClientId $ClientId

    Write-Host "Conectado."
    Write-Host ""

    # --------------------------------------------------------
    # BUSCAR ARQUIVOS
    # --------------------------------------------------------

    Write-Host "Buscando arquivos da pasta..."
    Write-Host ""

    $Itens = Get-PnPListItem `
        -List $Biblioteca `
        -FolderServerRelativeUrl $Pasta `
        -PageSize 500 `
        -Fields "FileLeafRef","FileRef","FSObjType"

    # --------------------------------------------------------
    # INVENTÁRIO DAS VERSÕES
    # --------------------------------------------------------

    foreach ($Item in $Itens) {

        # Ignora pastas
        if ($Item.FieldValues["FSObjType"] -ne 0) {
            continue
        }

        $TotalArquivos++

        $NomeArquivo =
            $Item.FieldValues["FileLeafRef"]

        $CaminhoArquivo =
            $Item.FieldValues["FileRef"]

        Write-Host "Analisando:"
        Write-Host $NomeArquivo

        try {

            $Versoes = @(
                Get-PnPFileVersion `
                    -Url $CaminhoArquivo
            )

            if ($Versoes.Count -eq 0) {

                Write-Host "  Nenhuma versão histórica."
                Write-Host ""

                continue
            }

            foreach ($Versao in $Versoes) {

                $TamanhoMB = [math]::Round(
                    ($Versao.Size / 1MB),
                    2
                )

                $Candidatos += [PSCustomObject]@{

                    Arquivo    = $NomeArquivo
                    Caminho    = $CaminhoArquivo
                    Versao     = $Versao.VersionLabel
                    IDVersao   = $Versao.Id
                    DataVersao = $Versao.Created
                    TamanhoMB  = $TamanhoMB
                }

                $TotalVersoes++
                $EspacoTotalMB += $TamanhoMB

                Write-Host `
                    "  Versão:" $Versao.VersionLabel `
                    "| ID:" $Versao.Id `
                    "| Tamanho:" $TamanhoMB "MB"
            }

            Write-Host ""
        }
        catch {

            Write-Host "  ERRO AO CONSULTAR:"
            Write-Host "  $($_.Exception.Message)"
            Write-Host ""

            $TotalErro++

            Add-Log `
                -Arquivo $NomeArquivo `
                -Caminho $CaminhoArquivo `
                -Versao "" `
                -IDVersao "" `
                -DataVersao "" `
                -TamanhoMB "" `
                -Resultado "ERRO" `
                -Mensagem $_.Exception.Message
        }
    }

    $EspacoTotalMB =
        [math]::Round($EspacoTotalMB, 2)

    # --------------------------------------------------------
    # RESUMO ANTES DA AÇÃO
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "=================================================="
    Write-Host " RESUMO DO INVENTÁRIO"
    Write-Host "=================================================="
    Write-Host ""

    Write-Host "Arquivos analisados:"
    Write-Host $TotalArquivos
    Write-Host ""

    Write-Host "Versões históricas encontradas:"
    Write-Host $TotalVersoes
    Write-Host ""

    Write-Host "Espaço aproximado:"
    Write-Host "$EspacoTotalMB MB"
    Write-Host ""

    if ($Candidatos.Count -eq 0) {

        Write-Host "Nenhuma versão histórica encontrada."
        Write-Host "Nada precisa ser feito."

        return
    }

    # --------------------------------------------------------
    # SIMULAÇÃO
    # --------------------------------------------------------

    if ($ModoExecucao -eq "SIMULACAO") {

        Write-Host "=================================================="
        Write-Host " MODO SIMULAÇÃO"
        Write-Host "=================================================="
        Write-Host ""

        Write-Host "NENHUMA versão será excluída."
        Write-Host ""

        foreach ($Candidato in $Candidatos) {

            Write-Host "Arquivo: $($Candidato.Arquivo)"
            Write-Host "Versão: $($Candidato.Versao)"
            Write-Host "ID: $($Candidato.IDVersao)"
            Write-Host "Ação: SERIA ENVIADA PARA A LIXEIRA"
            Write-Host ""

            Add-Log `
                -Arquivo $Candidato.Arquivo `
                -Caminho $Candidato.Caminho `
                -Versao $Candidato.Versao `
                -IDVersao $Candidato.IDVersao `
                -DataVersao $Candidato.DataVersao `
                -TamanhoMB $Candidato.TamanhoMB `
                -Resultado "SIMULACAO" `
                -Mensagem "Versão seria enviada para a Lixeira."
        }
    }

    # --------------------------------------------------------
    # EXECUÇÃO
    # --------------------------------------------------------

    if ($ModoExecucao -eq "EXECUCAO") {

        Write-Host "=================================================="
        Write-Host " ATENÇÃO - EXECUÇÃO REAL"
        Write-Host "=================================================="
        Write-Host ""

        Write-Host "Pasta:"
        Write-Host $Pasta
        Write-Host ""

        Write-Host "Versões que serão processadas:"
        Write-Host $Candidatos.Count
        Write-Host ""

        Write-Host "Espaço aproximado:"
        Write-Host "$EspacoTotalMB MB"
        Write-Host ""

        Write-Host "As versões históricas serão enviadas"
        Write-Host "para a Lixeira do SharePoint."
        Write-Host ""

        $Confirmacao =
            Read-Host "Digite EXECUTAR LIMPEZA para confirmar"

        if ($Confirmacao -ne "EXECUTAR LIMPEZA") {

            Write-Host ""
            Write-Host "Operação cancelada."
            Write-Host "Nenhuma versão foi excluída."

            return
        }

        Write-Host ""
        Write-Host "=================================================="
        Write-Host " INICIANDO LIMPEZA"
        Write-Host "=================================================="
        Write-Host ""

        foreach ($Candidato in $Candidatos) {

            Write-Host "Arquivo: $($Candidato.Arquivo)"
            Write-Host "Versão: $($Candidato.Versao)"
            Write-Host "ID: $($Candidato.IDVersao)"

            try {

                # ------------------------------------------------
                # REVALIDAÇÃO DE SEGURANÇA
                # ------------------------------------------------

                $VersoesAtuais = @(
                    Get-PnPFileVersion `
                        -Url $Candidato.Caminho
                )

                $VersaoAindaExiste =
                    $VersoesAtuais |
                    Where-Object {
                        $_.Id -eq $Candidato.IDVersao
                    }

                if ($null -eq $VersaoAindaExiste) {

                    Write-Host "Resultado: IGNORADO"
                    Write-Host "A versão não existe mais."
                    Write-Host ""

                    $TotalIgnorado++

                    Add-Log `
                        -Arquivo $Candidato.Arquivo `
                        -Caminho $Candidato.Caminho `
                        -Versao $Candidato.Versao `
                        -IDVersao $Candidato.IDVersao `
                        -DataVersao $Candidato.DataVersao `
                        -TamanhoMB $Candidato.TamanhoMB `
                        -Resultado "IGNORADO" `
                        -Mensagem "Versão não encontrada na revalidação."

                    continue
                }

                # ------------------------------------------------
                # EXCLUSÃO
                # ------------------------------------------------

                Remove-PnPFileVersion `
                    -Url $Candidato.Caminho `
                    -Identity $Candidato.IDVersao `
                    -Recycle `
                    -Force

                Write-Host "Resultado: SUCESSO"
                Write-Host ""

                $TotalSucesso++

                Add-Log `
                    -Arquivo $Candidato.Arquivo `
                    -Caminho $Candidato.Caminho `
                    -Versao $Candidato.Versao `
                    -IDVersao $Candidato.IDVersao `
                    -DataVersao $Candidato.DataVersao `
                    -TamanhoMB $Candidato.TamanhoMB `
                    -Resultado "SUCESSO" `
                    -Mensagem "Versão enviada para a Lixeira."
            }
            catch {

                Write-Host "Resultado: ERRO"
                Write-Host $_.Exception.Message
                Write-Host ""

                $TotalErro++

                Add-Log `
                    -Arquivo $Candidato.Arquivo `
                    -Caminho $Candidato.Caminho `
                    -Versao $Candidato.Versao `
                    -IDVersao $Candidato.IDVersao `
                    -DataVersao $Candidato.DataVersao `
                    -TamanhoMB $Candidato.TamanhoMB `
                    -Resultado "ERRO" `
                    -Mensagem $_.Exception.Message
            }
        }
    }

    # --------------------------------------------------------
    # RESUMO FINAL
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "=================================================="
    Write-Host " RESUMO FINAL"
    Write-Host "=================================================="
    Write-Host ""

    Write-Host "Modo:"
    Write-Host $ModoExecucao
    Write-Host ""

    Write-Host "Arquivos analisados:"
    Write-Host $TotalArquivos
    Write-Host ""

    Write-Host "Versões encontradas:"
    Write-Host $TotalVersoes
    Write-Host ""

    if ($ModoExecucao -eq "EXECUCAO") {

        Write-Host "Sucessos:"
        Write-Host $TotalSucesso
        Write-Host ""

        Write-Host "Ignorados:"
        Write-Host $TotalIgnorado
        Write-Host ""

        Write-Host "Erros:"
        Write-Host $TotalErro
        Write-Host ""
    }

    Write-Host "Espaço processado aproximadamente:"
    Write-Host "$EspacoTotalMB MB"
}
catch {

    Write-Host ""
    Write-Host "=================================================="
    Write-Host " ERRO GERAL"
    Write-Host "=================================================="
    Write-Host ""

    Write-Host $_.Exception.Message
}
finally {

    # --------------------------------------------------------
    # EXPORTAR LOG
    # --------------------------------------------------------

    if ($Log.Count -gt 0) {

        $Log |
            Export-Csv `
                -Path $ArquivoLog `
                -NoTypeInformation `
                -Delimiter ";" `
                -Encoding utf8BOM

        Write-Host ""
        Write-Host "=================================================="
        Write-Host " LOG GERADO"
        Write-Host "=================================================="
        Write-Host ""

        Write-Host $ArquivoLog
    }

    Disconnect-PnPOnline `
        -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "Fim da execução."
}
