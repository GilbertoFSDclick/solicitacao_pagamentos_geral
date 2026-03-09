# Análise Completa: Projeto Atual vs. Documento de Requisitos

## Sumário Executivo

O projeto foi **adaptado** para o processo **Lançamento de Solicitações de Pagamento Geral** conforme o documento **Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral**. O fluxo utiliza **Admin → NBS Fiscal → Entrada Só Diversa** (sem XML/SAAM), planilha TIPOS_DE_PAGAMENTO para código contábil, e exceções TE01–TE08 implementadas.

**Status:** Projeto alinhado ao documento. Componentes do fluxo antigo (Despesas de Produtos/NF-e) permanecem em `modulos/nbs/__setup.py` para referência, mas o fluxo principal usa `solicitacao_pagamento.py`.

---

## 1. Estrutura do Projeto Atual

### 1.1 Arquitetura Geral

```
original-rpa-produtos-nbs-main/
├── main.py                 # Fluxo principal
├── operacoes/
│   └── tratar_tarefa.py    # Extração de dados do Holmes
├── src/
│   ├── webhook.py          # Obtenção de processos via webhook
│   └── exceptions.py       # Exceções de negócio/técnicas
├── modulos/
│   ├── nbs/                # Interação com NBS (entrada de NF-e)
│   ├── holmes/             # API Holmes (tarefas, processos)
│   ├── saam/               # API SAAM (obtenção de XML)
│   ├── webhook/            # Webhook + integração Holmes
│   ├── interface/          # Elementos de UI (pywinauto)
│   ├── dashboard/          # Dashboard
│   └── ambiente/           # Configurações de ambiente
├── .ini                    # Configurações
├── requirements.txt
└── imagens/                # Imagens para reconhecimento (bot.imagem)
```

### 1.2 Fluxo Atual (Despesas de Produtos)

1. **Webhook** → Obtém processos com filtros: Torre=UAB, Tipo Documento=Despesas Produtos/DANFE NBS, DMS=NBS
2. **Holmes** → Consulta tarefa e processo para extrair: Filial, Número NF, Valor, CNPJ Emissor, Centro de Custo, Classificação Contábil, Torre, Data Vencimento, Quantidade Parcelas, Protocolo, etc.
3. **SAAM** → Obtém XML da NF-e (por chave ou por CNPJ+numero+data)
4. **NBS** → Fluxo completo:
   - Login → Contas a Pagar → Nota Fiscal de Compra → **NF-e**
   - Monitor NF-e → Pesquisar nota → Aceitar
   - Capa (protocolo, PIS/COFINS)
   - **CFOPs** (leitura do XML)
   - **Contabilização** (Classificação Contábil, Centro de Custo, Histórico Padrão)
   - **Faturamento** (parcelas, Tipo Pagamento, Natureza Despesa, Vencimento)
5. **Holmes** → Encaminhar tarefa (sucesso/erro)
6. **Email** → Notificação com log

### 1.3 Dados Utilizados (Projeto Atual)

| Campo Holmes | Uso |
|--------------|-----|
| Filial | De/para Empresa/Filial NBS |
| Número | Número da NF |
| Valor | Validação |
| CNPJ Emissor | Busca no Monitor NF-e |
| Centro de Custo | Contabilização |
| Classificação Contábil | Contabilização |
| Torre | Histórico Padrão, Tipo Pagamento, Natureza Despesa |
| Data de Vencimento | Faturamento |
| Quantidade de Parcelas | Faturamento |
| Protocolo | Campo Observações na Capa |
| Tipo de Documento | Despesas Produtos vs Despesas Serviços (PIS/COFINS) |

### 1.4 Planilhas/De-Para Atuais

- **Empresas_NBS.xlsx** → Empresa/Filial por CNPJ
- **Empresas_NBS - Despesas.xlsx** → Filtro de CNPJs
- **CONTAS CTB - CREDITO PISCOFINS AUTOMOB.xlsx** → Validação PIS/COFINS
- **CFOP** → Hardcoded em `de_para_cfop()` no código

---

## 2. Requisitos do Documento (Solicitação de Pagamento Geral)

### 2.1 Objetivo

> Automatizar o processo de extração de informações de processo do Holmes para lançamento de dados no Dealer (NBS).

### 2.2 Premissas

- Acesso à planilha **"Solicitação de Pagamento Geral"**
- Processo aprovado pelo Gestor no Holmes
- Webhook aciona a automação

### 2.3 Fluxo do Documento

| Etapa | Descrição |
|-------|-----------|
| 1 | Webhook aciona automação (ID Fluxo a definir) |
| 2 | Leitura de Nota anexada ao processo (ferramenta) |
| 3 | **Consulta planilha TIPOS_DE_PAGAMENTO.XLS** → Coluna A (Tipo Pagamento Holmes) → Coluna D (CODIGO_UAB_CONTABIL) |
| 4 | Inclusão em banco de controle (dados a definir) |
| 5 | **Lançamento no NBS** (fluxo principal) |
| 6 | Remoção tarefa webhook + encaminhamento |
| 7 | Email de sucesso/erro |

### 2.4 Fluxo NBS (Documento) – Diferente do Atual

| Passo | Ação |
|-------|------|
| 1 | Abrir NBS, login (USUARIO_NBS, SENHA_NBS) |
| 2 | **Aba Admin** → **NBS Fiscal** (não Contas a Pagar) |
| 3 | Empresa, Filial, Exercício Contábil (Código Empresa/Filial do Holmes) |
| 4 | **Entrada** → **Incluir Entrada (Só Diversa)** |
| 5 | **Fornecedor CNPJ/CPF** ← campo CPF do Holmes (TE01: pendenciar se não encontrado) |
| 6 | Número NF ← Processo Holmes; Série = "E"; Emissão/Entrada ← "Iniciar em" Holmes |
| 7 | Desmarcar "Quero esta nota no livro fiscal" |
| 8 | Total Nota ← Despesa Pagamento Holmes; Observações ← Protocolo Holmes |
| 9 | Contabilização: Conta Contábil (código UAB da planilha); Centro de Custo ← Centro de Custo UAB NBS Holmes |
| 10 | Faturamento: Total Parcelas = 1 |
| 11 | Tipo Pagamento = Boleto; Natureza Despesa = Outras Despesas |
| 12 | Vencimento ← Data de Vencimento Holmes (validar Data de Corte) |
| 13 | Confirmar → Guardar num_controle → Fechar Ficha de Controle (Cancelar) |

### 2.5 Campos Holmes (Documento)

| Campo Holmes | Uso no NBS |
|--------------|------------|
| Código Empresa NBS | Empresa |
| Código Filial NBS | Filial |
| CPF | Fornecedor CNPJ/CPF |
| Processo | Número da NF |
| Iniciar em | Emissão e Entrada |
| Despesa Pagamento | Total Nota |
| Protocolo | Observações |
| Centro de Custo - UAB - NBS | Centro de Custo |
| Data de Vencimento | Vencimento |
| Tipo de Pagamento | Consulta planilha → Código Contabilização |

### 2.6 Exceções (Documento)

| ID | Exceção | Tratamento |
|----|---------|------------|
| TE01 | CNPJ não encontrado | Pendenciar Holmes |
| TE02 | Problemas com NBS | Email + log |
| TE03 | Travamento de Interface | Email + log |
| TE04 | Problemas WebHook/Holmes | Email + log; alocar "Lançamento Manual RPA falhou" |
| TE05 | Erro no processamento | Pendenciar; alocar "Lançamento Manual RPA falhou" |
| TE06 | Falta de memória/disco | Pendenciar |
| TE07 | Tipo Pagamento não na planilha | Não lançar; Business Exception; comentário Holmes; email Fiscal |
| TE08 | Holmes não carrega | Retry 3x; System Exception |

---

## 3. Comparativo: Projeto Atual vs. Requisitos

| Aspecto | Projeto Atual | Documento de Requisitos |
|---------|---------------|--------------------------|
| **Tipo de processo** | Despesas de Produtos (NF-e) | Solicitação de Pagamento Geral |
| **Origem dos dados** | Holmes + XML (SAAM) | Holmes + planilha TIPOS_DE_PAGAMENTO |
| **Entrada no NBS** | Contas a Pagar → NF-e (Monitor) | Admin → NBS Fiscal → Entrada Só Diversa |
| **Identificação fornecedor** | CNPJ Emissor (XML) | CPF (Holmes) |
| **Série** | Do XML | Fixa "E" |
| **CFOPs** | Do XML (produtos) | Não aplicável |
| **Contabilização** | Classificação Contábil Holmes + de/para | Tipo Pagamento → planilha → Código UAB |
| **Parcelas** | Múltiplas (Quantidade Parcelas) | Sempre 1 |
| **Tipo/Natureza** | Por Torre (de/para no código) | Boleto / Outras Despesas (fixo) |
| **Planilhas** | Empresas_NBS, CONTAS CTB, Despesas | TIPOS_DE_PAGAMENTO.XLS |
| **SAAM** | Sim (XML) | Não (sem XML) |
| **Livro fiscal** | Sim (NF-e) | Não (desmarcar flag) |

---

## 4. Plano de Adaptação

### 4.1 Componentes Reutilizáveis

- **Webhook** (`src/webhook.py`, `modulos/webhook/`) – estrutura mantida; ajustar filtros e `Properties`
- **Holmes** (`modulos/holmes/`) – APIs reutilizáveis
- **NBS** – login e estrutura base; **fluxo de entrada totalmente novo**
- **Email** – mantido
- **Configuração** (`.ini`) – novos parâmetros

### 4.2 Componentes a Criar/Substituir

| Componente | Ação |
|------------|------|
| `operacoes/tratar_tarefa.py` | Reescrever para campos do documento (CPF, Código Empresa/Filial, Despesa Pagamento, Protocolo, Centro de Custo, Data Vencimento, Tipo Pagamento) |
| `src/webhook.py` | Novo `Properties`; novos filtros (fluxo Solicitação Pagamento Geral) |
| `modulos/nbs/__setup.py` | Novo fluxo: Admin → NBS Fiscal → Entrada Só Diversa (sem NF-e, sem CFOPs, sem SAAM) |
| Planilha TIPOS_DE_PAGAMENTO | Implementar leitura: Coluna A → Coluna D (CODIGO_UAB_CONTABIL) |
| Banco de controle | Implementar se exigido (dados a definir no documento) |

### 4.3 Componentes a Remover/Desativar

- **SAAM** – não usado no novo fluxo
- **Leitura de XML** (`produtos_xml`, `valores_xml`, `de_para_cfop`)
- **Fluxo NF-e** (Monitor, Aceitar, CFOPs, PIS/COFINS)
- **Imagens** específicas de NF-e (manter apenas as do novo fluxo)

### 4.4 Novos Parâmetros (.ini)

```
[nbs]
# Existentes: usuario, senha, servidor, caminho_sistema, titulo_sistema
planilha_tipos_pagamento = <caminho>/TIPOS_DE_PAGAMENTO.XLS
diretorio_planilha = <DIRETORIO_PLANILHA>

[holmes]
id_fluxo = <a definir>
nome_processo = Lançamento de Solicitações de Pagamento
# ... ajustes conforme novo fluxo
```

### 4.5 Tratamento de Exceções

Implementar TE01–TE08 conforme tabela do documento, incluindo:

- TE01: pendenciar com motivo "CNPJ não encontrado"
- TE07: não lançar; Business Exception; comentário no Holmes; email para área Fiscal
- TE08: retry 3x antes de System Exception

---

## 5. Riscos e Dependências

1. **Definições pendentes no documento**
   - ID Fluxo Holmes
   - Nome do fluxo
   - Lista completa de campos do Holmes
   - Estrutura do banco de controle
   - Data de Corte para validação de vencimento

2. **Planilha TIPOS_DE_PAGAMENTO**
   - Formato exato (colunas, abas)
   - Localização e permissões de acesso

3. **Interface NBS**
   - Tela "Admin → NBS Fiscal → Entrada Só Diversa" pode exigir novos seletores (pywinauto) e imagens para o bot

4. **Webhook**
   - Novo fluxo e filtros a configurar no serviço de webhook

---

## 6. Próximos Passos Recomendados

1. Obter do negócio:
   - ID e nome do fluxo Holmes para Solicitação de Pagamento Geral
   - Planilha TIPOS_DE_PAGAMENTO (exemplo e caminho)
   - Definição da Data de Corte
   - Lista final de campos do Holmes

2. Criar branch de desenvolvimento para a adaptação.

3. Implementar em fases:
   - Fase 1: Novo `tratar_tarefa` + `Properties` + filtros webhook
   - Fase 2: Leitura da planilha TIPOS_DE_PAGAMENTO
   - Fase 3: Novo fluxo NBS (Admin → NBS Fiscal → Entrada Só Diversa)
   - Fase 4: Exceções TE01–TE08
   - Fase 5: Banco de controle (se aplicável) e ajustes de email

4. Validar em ambiente de homologação com processos reais do novo fluxo.
