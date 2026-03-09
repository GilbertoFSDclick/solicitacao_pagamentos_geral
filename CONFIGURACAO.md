# Configuração - Lançamento de Solicitações de Pagamento Geral

## Ambiente (venv)

1. **Criar venv e instalar dependências:** execute `.\setup.ps1` na pasta do projeto.
2. **Rodar o bot:** execute `.\start.ps1`. O script usa o Python do venv quando existir; caso contrário, usa o Python do PATH.

## Parâmetros a definir (conforme documento)

Atualize o arquivo `.ini` quando os valores forem definidos:

- `[holmes] id_fluxo` - ID do fluxo Holmes para Solicitação de Pagamento Geral
- `[holmes] nome_processo` - Nome do processo (ex: "Lançamento de Solicitações de Pagamento")
- `[holmes] nome_atividade` - Nome da atividade/tarefa que a RPA assume (ex: "Lançamento RPA")
- `[nbs] planilha_tipos_pagamento` - Caminho da planilha TIPOS_DE_PAGAMENTO (Col A = Tipo Pagamento, Col D = Código UAB)

## Campos do Holmes (processo Solicitação de Pagamentos Geral - real)

Mapeamento dos campos que constam nos detalhes da tarefa:

| Campo Holmes | Uso no fluxo |
|--------------|--------------|
| Protocolo | Identificador, Observações |
| Filial | Nome da filial (ex: 481 - GWM GUARULHOS STA FRANCI) |
| CPF | Fornecedor CNPJ/CPF |
| Tipo de pagamento | Consulta planilha -> Código UAB (ex: Multa à cobrar) |
| Valores / Devolução de adiantamento/Obrigação | Total Nota (Despesa Pagamento) |
| Observação | Observações na NF |
| CNPJ da Loja | De/para Empresas_NBS -> Código Empresa/Filial NBS |
| Torre | Aba da planilha Empresas_NBS (UAB se for ID) |
| AIT | Número da NF (Processo) |
| Data de vencimento | Vencimento (usa data atual se ausente) |
| Centro de Custo | Contabilização (usa 0 se ausente) |

Código Empresa/Filial NBS: obtido via planilha Empresas_NBS usando CNPJ da Loja + Torre.

## Planilha TIPOS_DE_PAGAMENTO

- **Aba:** Plan1 (ou Planilha1)
- **Header:** Linha 1 (Tipos de Pagamentos | Autostar | Original SP | UAB)
- **Coluna A:** Tipos de Pagamentos (valor do Holmes, ex: "Multa à cobrar")
- **Coluna D:** Código UAB (CODIGO_UAB_CONTABIL)

Formato: .xlsx. Caminho em `planilha_tipos_pagamento` no .ini (ex: `..\TIPOS_DE_PAGAMENTO.xlsx` para teste local).

## Webhook

O webhook deve retornar processos com as propriedades acima. Os nomes das propriedades são normalizados (minúsculas, sem acentos). Ajuste `Properties` em `src/webhook.py` se os nomes do Holmes forem diferentes.

## Fluxo NBS - Ajustes possíveis

O fluxo em `modulos/nbs/solicitacao_pagamento.py` navega: Admin → NBS Fiscal → Entrada → Incluir Entrada (Só Diversa). Os seletores (class_name, título de janela) podem precisar de ajuste conforme a versão do NBS. Teste em homologação e ajuste os elementos de interface conforme necessário.
