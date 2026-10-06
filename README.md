# Automação de versões do SharePoint

Projeto em PowerShell para estudar e automatizar a consulta e a limpeza controlada do histórico de versões de arquivos.

## Arquivos

| Script | Função |
| --- | --- |
| `01-Inventario-Versoes.ps1` | Estrutura inicial de estudo. Não conecta ao SharePoint, não coleta versões e não exclui nada. |
| `02-Relatorio-Versoes.ps1` | Conecta ao SharePoint e gera um CSV com as versões consultadas. Não exclui versões. |
| `06-Limpeza-Versoes-V1.ps1` | Consulta versões e registra candidatos. Começa em `SIMULACAO`, sem exclusões. |

## Preparação

Use PowerShell 7, o módulo PnP.PowerShell e um aplicativo Microsoft Entra ID configurado para autenticação interativa, com as permissões e os consentimentos necessários. A conta também precisa ter acesso ao escopo escolhido.

Em uma cópia local dos scripts 02 e 06, configure:

- `SiteUrl`: endereço do seu site.
- `Biblioteca`: título da biblioteca.
- `Pasta`: caminho relativo ao servidor da pasta escolhida. Use o caminho real; ele pode ser diferente do título exibido na biblioteca.
- `ClientId`: identificador do aplicativo usado na conexão.

O ClientId não é uma senha, mas foi removido desta cópia para não expor a configuração interna do ambiente de origem. Não adicione senhas, tokens ou certificados ao repositório.

Os scripts 02 e 06 interrompem a execução enquanto o endereço de exemplo ou o ClientId vazio permanecerem configurados. Os relatórios ficam na pasta `Relatorios`, junto aos scripts.

## Ordem de uso

1. Configure o script 02 e execute-o para consultar o escopo escolhido.
2. Revise o CSV gerado. Linhas com erros precisam ser analisadas.
3. Configure o mesmo escopo no script 06 e mantenha `SIMULACAO`.
4. Execute a simulação e confira os candidatos e o log.

O script 01 é material de estudo, não um pré-requisito para executar os demais.

## Execução real

O script 06 possui um modo `EXECUCAO`, com pasta obrigatória, confirmação digitada e reconsulta da versão antes da remoção. A chamada de remoção existente usa `Remove-PnPFileVersion` com `-Recycle` e `-Force`.

Antes de usar esse modo, valide o comportamento e a compatibilidade dos comandos com o módulo instalado em um ambiente de teste. Revise também as regras de retenção aplicáveis. O relatório de candidatos não autoriza a exclusão e o tamanho somado das versões não comprova espaço efetivamente liberado.

## Publicação e revisão

Estas são cópias dos scripts originais com dados de configuração substituídos por exemplos. A lógica original foi preservada, com uma verificação adicional para impedir conexão com valores não configurados.

Foi feita revisão textual dos arquivos. Estas cópias não foram executadas contra o SharePoint e não tiveram a sintaxe validada por um interpretador PowerShell nesta revisão.

Não publique CSVs, logs ou capturas de tela que contenham informações internas. O `.gitignore` auxilia futuros envios com Git; no envio manual pelo navegador, selecione os arquivos explicitamente.
