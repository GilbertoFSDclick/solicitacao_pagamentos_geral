# Checklist - Interface NBS (Solicitação Pagamento Geral)

Quando o NBS for atualizado, verificar se os elementos abaixo ainda existem. O fluxo usa **Admin → NBS Fiscal → Entrada Só Diversa**.

## Janelas (títulos esperados)

| Etapa | Título da janela |
|-------|------------------|
| Login | NBS Sistema Financeiro |
| Pós-login | Sistema Financeiro - SISFIN |
| Entrada | Contém "Entrada" ou "Diversa" |
| Contabilização | Contém "Incluir", "Conta" ou "Contabilização" |
| Pop-up sucesso | Contém "Aviso" ou "Informação" |
| Ficha Controle | Contém "Ficha" ou "Controle" |

## Elementos pywinauto (class_name)

| Classe | Uso |
|--------|-----|
| `TOvcPictureField` | Empresa, Filial, campo SisFin |
| `TMenuItem` | Menu Admin → NBS Fiscal |
| `TEdit` | Fornecedor CPF, Número NF, Conta Contábil, Centro de Custo |
| `TTabSheet` | Abas Contabilização, Faturamento |
| `TBitBtn` | Botões +, Confirmar, OK, Gerar |
| `TOvcNumericField` | Total Parcelas |
| `TwwDBLookupCombo` | Tipo Pagamento, Natureza Despesa |
| `TOvcDbPictureField` | Data de Vencimento |

## Fluxo de navegação

1. **Alt+A** → Admin
2. **NBS Fiscal** (menu)
3. **Alt+E** → Entrada
4. **Down, Down, Enter** → Incluir Entrada (Só Diversa)

## TE01 - Detecção CNPJ não encontrado

Janelas com "não encontrado", "não cadastrado" (ou variações) indicam que o CPF/CNPJ não foi localizado.

## Arquivo de referência

`modulos/nbs/solicitacao_pagamento.py` – função `processar_entrada_solicitacao_pagamento`
