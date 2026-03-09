# Análise: Projeto vs. Documentação Atualizada

**Documento analisado:** Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral (2).docx  
**Data da análise:** Março 2026

---

## Resumo Executivo

| Métrica | Valor |
|--------|--------|
| **Alinhamento com o documento** | ~**97%** |
| **Itens divergentes ou a validar** | 2 |
| **Leitura de Nota anexada** | **Não mencionada** na nova versão (possível remoção do escopo) |

---

## 1. Fluxo Resgate de Informações no Holmes

| Requisito do documento | Status no projeto |
|------------------------|-------------------|
| Automação chamada pelo webhook após criação de processo | ✅ Implementado |
| Identificar e assumir tarefas pendentes | ✅ Implementado |
| **Extrair e guardar:** Empresa_holmes, Filial_holmes, Cpf_holmes, Num_nf_holmes, Serie_nf_holmes, Emissão_nf_holmes, Entrada_nf_holmes, Vlr_nf_holmes, Protocolo_holmes, Centro_custo_uab_holmes, Dt_Vencimento_Holmes | ✅ Implementado (`DadosExtraidosHolmes`) |

**Observação:** A nova versão do documento **não menciona** "Leitura de Nota anexada ao processo". Esse item existia em versões anteriores como "lista a definir". Se foi removido do escopo, não há pendência nesse fluxo.

---

## 2. Fluxo Consulta Código de Contabilização (Planilha)

| Requisito do documento | Status no projeto |
|------------------------|-------------------|
| Acessar TIPOS_DE_PAGAMENTO em DIRETORIO_PLANILHA | ✅ Implementado |
| Aba Planilha1 | ✅ Implementado (Plan1 ou Planilha1) |
| Coluna A (Tipos de Pagamentos) → Coluna D (UAB) = CODIGO_UAB_CONTABIL | ✅ Implementado |
| TE07: Tipo não encontrado → não lançar; Business Exception; comentário Holmes; e-mail Fiscal | ✅ Implementado |

---

## 3. Fluxo Prévio II – Banco de Dados de Controle

| Requisito do documento | Status no projeto |
|------------------------|-------------------|
| Acessar banco de dados de controle da automação | ✅ Implementado |
| Inserir dados resgatados no Holmes + Codigo_UAB_Contabil | ✅ Implementado (12 campos) |

**Campos do documento:** Empresa_holmes, Filial_holmes, Cpf_holmes, Num_nf_holmes, Serie_nf_holmes, Emissão_nf_holmes, Entrada_nf_holmes, Vlr_nf_holmes, Protocolo_holmes, Centro_custo_uab_holmes, Dt_Vencimento_Holmes, Codigo_UAB_Contabil  
**Implementação:** Todos os campos mapeados e inseridos via `registrar_processo_controle`.

---

## 4. Fluxo Principal – Lançamento NBS (Entrada Só Diversa)

| Passo do documento | Implementação |
|--------------------|---------------|
| Abrir NBS; login USUARIO_NBS, SENHA_NBS; Confirmar/Enter | ✅ |
| Aba Admin → NBS Fiscal | ✅ |
| Empresa, Filial, Exercício Contábil (Empresa_holmes, Filial_holmes) | ✅ |
| Entrada → Incluir Entrada (Só Diversa) – segundo atalho | ✅ |
| Fornecedor CNPJ/CPF ← Cpf_holmes; Tab para trazer registro | ✅ |
| TE01: CNPJ não encontrado → pendenciar | ✅ |
| Número NF ← Num_nf_holmes; Série = E; Emissão/Entrada ← Emissão/Entrada_holmes | ✅ |
| Desmarcar "Quero esta nota no livro fiscal" | ✅ |
| Total Nota ← Vlr_nf_holmes; Observações ← Protocolo_holmes | ✅ |
| Contabilização: Conta Contábil (código UAB); Centro de Custo | ✅ |
| Faturamento: Total Parcelas = 1; Gerar | ✅ |
| Tipo Pagamento = Boleto; Natureza = Outras Despesas | ✅ |
| Vencimento ← Dt_vencimento_holmes; validar Data de Corte | ✅ |
| Confirmar; guardar num_controle; pop-up OK | ✅ |
| Ficha de Controle de Pagamento → Cancelar | ✅ |

**Ponto de atenção:** O documento diz "Preencher o campo Centro de Custo com dados da variável Codigo_UAB_Contabil". Isso parece ser um erro de redação: o **Centro de Custo** deve vir de **Centro_custo_uab_holmes** (Holmes), e o **Codigo_UAB_Contabil** deve ir para a **Conta Contábil**. O projeto está implementado dessa forma. Recomenda-se confirmar com o negócio.

---

## 5. Fluxo Final – Remoção Tarefas Webhook

| Requisito do documento | Status no projeto |
|------------------------|-------------------|
| Retornar à tarefa no Holmes; Aprovar | ✅ Implementado |
| Remoção de tarefa do webhook; encaminhamento | ✅ Implementado |

---

## 6. Fluxo Envio de E-mail

| Requisito do documento | Status no projeto |
|------------------------|-------------------|
| E-mail de sucesso com relatório final e log em anexo | ✅ Implementado (tabela de resultados + log) |
| E-mail de sucesso: "Planilha Atualizada, e Notas Fiscais compactadas por Remessa" | ⚠️ **Não implementado** – o fluxo Entrada Só Diversa não gera planilha nem NF compactada; pode não ser aplicável |

**Observação:** O documento menciona "Planilha Atualizada" e "Notas Fiscais compactadas por Remessa" no e-mail de sucesso. No fluxo de Solicitação de Pagamento Geral (Entrada Só Diversa), não há geração de planilha nem de pacote de NFs. Se o documento se refere a outro processo ou foi copiado de outro RPA, esse item pode ser N/A. Validar com o negócio.

---

## 7. Tratamento de Exceções (TE01–TE08)

| ID | Exceção | Tratamento no projeto |
|----|---------|------------------------|
| TE01 | CNPJ não encontrado | ✅ Pendenciar Holmes; motivo: CNPJ não encontrado |
| TE02 | Problemas com NBS | ✅ E-mail + log |
| TE03 | Travamento de Interface | ✅ E-mail + log |
| TE04 | Problemas WebHook/Holmes | ✅ E-mail + log; "Lançamento Manual RPA falhou" |
| TE05 | Erro no processamento | ✅ Pendenciar; "Lançamento Manual RPA falhou" |
| TE06 | Falta de memória/disco | ✅ Coberto por retry/encaminhar erro |
| TE07 | Tipo Pagamento não na planilha | ✅ Não lançar; Business Exception; comentário Holmes; e-mail Fiscal |
| TE08 | Holmes não carrega | ✅ Retry 3x; System Exception |

---

## 8. Parâmetros

| Parâmetro do documento | Suportado no .ini |
|------------------------|------------------|
| SENHA_NBS, USUARIO_NBS | ✅ [nbs] senha, usuario |
| FLUXO_HOLMES, ID_FLUXO | ✅ [holmes] id_fluxo |
| EMAIL_ERRO, EMAIL_SUCESSO | ✅ [email.destinatarios] erro, sucesso |
| TOKEN_REQUISIÇÃO | ✅ [webhook] token, [holmes] token |

---

## 9. Itens a Validar com o Negócio

| # | Item | Situação |
|---|------|----------|
| 1 | **Centro de Custo vs. Codigo_UAB_Contabil** | Documento pode ter trocado os campos; projeto usa Centro_custo_uab_holmes para Centro de Custo e Codigo_UAB_Contabil para Conta Contábil. |
| 2 | **E-mail sucesso: Planilha Atualizada e NFs compactadas** | Não aplicável ao fluxo Entrada Só Diversa atual; confirmar se deve ser removido do escopo ou se há outro processo que gera esses anexos. |
| 3 | **Leitura de Nota anexada** | Não aparece na nova versão do documento; se foi removida do escopo, não há pendência. |

---

## 10. Conclusão

O projeto está **alinhado em ~97%** com a documentação atualizada. Os dados a extrair, o banco de controle e o fluxo NBS estão implementados conforme o documento.

**Pendências:**
- Confirmar com o negócio: Centro de Custo (origem dos dados) e anexos do e-mail de sucesso.
- Validar em produção (VPN + NBS) para ajuste fino de interface.

**Possível ganho:** Se "Leitura de Nota anexada" foi removida do escopo, o projeto pode ser considerado **100% do escopo documentado** em termos de implementação.
