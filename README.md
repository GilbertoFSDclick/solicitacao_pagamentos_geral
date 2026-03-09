# Automação - Lançamento de Solicitações de Pagamento Geral (NBS)

Conforme **Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral**.

## Objetivo

Automatizar o processo de extração de informações de processo do Holmes para lançamento de dados no NBS (Dealer).

## Premissas

- Acesso ao Holmes
- Acesso ao NBS
- Planilha TIPOS_DE_PAGAMENTO (em `planilha_tipos_pagamento` no .ini)
- Processo aprovado pelo Gestor no Holmes
- Automação acionada pelo webhook após criação de processo

## Fluxo

### 1. Webhook → Obtenção de processos
- Automação chamada pelo webhook após criação de processo no Holmes
- Filtros: DMS = NBS

### 2. Holmes → Extração de dados
- CPF, Protocolo, Tipo de Pagamento, Valores/Despesa Pagamento
- Código Empresa NBS, Código Filial NBS (prioridade sobre planilha)
- Processo (Número NF), Iniciar em, Centro de Custo - UAB - NBS
- Data de Vencimento

### 3. Planilha TIPOS_DE_PAGAMENTO
- Acessar planilha em `planilha_tipos_pagamento`
- Aba Plan1 ou Planilha1
- Coluna A (Tipos de Pagamentos) → Coluna D (UAB) = CODIGO_UAB_CONTABIL

### 4. NBS → Lançamento
- Admin → NBS Fiscal → Entrada → Incluir Entrada (Só Diversa)
- Fornecedor CNPJ/CPF ← CPF Holmes
- Número NF ← Processo Holmes | Série = E | Emissão/Entrada ← Iniciar em
- Desmarcar "Quero esta nota no livro fiscal"
- Total Nota ← Despesa Pagamento | Observações ← Protocolo
- Contabilização: Conta Contábil (código UAB) | Centro de Custo ← Holmes
- Faturamento: Total Parcelas = 1 | Tipo = Boleto | Natureza = Outras Despesas
- Vencimento ← Data de Vencimento (validar Data de Corte se configurada)

### 5. Holmes → Encaminhamento
- Sucesso: Finalizar tarefa
- Erro: Pendenciar com motivo

### 6. E-mail
- Sucesso: lista de sucesso com relatório
- Erro: lista de erro
- TE07 (Tipo não parametrizado): notificação para área Fiscal

## Exceções (TE01–TE08)

| ID | Exceção | Tratamento |
|----|---------|------------|
| TE01 | CNPJ não encontrado | Pendenciar Holmes |
| TE02 | Problemas com NBS | E-mail + log |
| TE03 | Travamento de Interface | E-mail + log |
| TE04 | WebHook/Holmes | E-mail + log; alocar "Lançamento Manual RPA falhou" |
| TE05 | Erro no processamento | Pendenciar; alocar "Lançamento Manual RPA falhou" |
| TE06 | Falta de memória/disco | Pendenciar |
| TE07 | Tipo Pagamento não na planilha | Não lançar; comentário Holmes; e-mail Fiscal |
| TE08 | Holmes não carrega | Retry 3x; System Exception |

## Configuração

Copie `.ini.example` para `.ini` e preencha os parâmetros. Consulte `INI_GUIA.md` e `CONFIGURACAO.md`.
