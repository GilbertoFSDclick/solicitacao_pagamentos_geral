# O que fazer para entregar 100% do projeto

Guia objetivo para fechar os ~8% restantes e considerar o projeto entregue.

---

## Visão geral

| Fase | O que é | Quem define/faz |
|------|---------|-----------------|
| **A** | Itens “a definir” no documento | Negócio/gestor |
| **B** | Configuração e ambiente | Você + infra/cliente |
| **C** | Teste e validação no NBS | Você (na VPN/máquina do cliente) |
| **D** | Ajustes pós-teste e entrega | Você |

---

## Fase A – Fechar itens “a definir” (documento)

O documento deixa dois pontos em aberto. Sem isso, 100% do **escopo documentado** não fica fechado.

### A.1 Leitura de Nota anexada ao processo

- **Documento:** “Automação deverá acionar ferramenta para leitura de Nota anexada ao processo. Aguardar dados – Lista completa de dados será definida.”
- **O que fazer:**
  1. Pedir ao negócio/gestor: **lista de dados** a extrair da nota anexada (quais campos, de qual anexo, formato).
  2. Se for obrigatório: implementar a leitura (API Holmes de documentos da tarefa já existe: `consulta_documento_tarefa`) e mapear os campos na automação.
  3. Se o negócio disser que **não será usado** ou “manter como está”: registrar por escrito e considerar requisito atendido (sem desenvolvimento).

### A.2 Banco de dados de controle

- **Documento:** “Inserir dados resgatados no Holmes. Dados a definir.”
- **O que fazer:**
  1. Pedir ao negócio: **modelo do banco** (tabela(s), colunas) e **quem provê** (conexão, string, ambiente).
  2. Com o modelo definido: implementar o insert em `operacoes/banco_controle.py` (a função `registrar_processo_controle` já é chamada; falta só o insert real).
  3. Se o negócio disser que **não haverá banco de controle**: remover ou desativar a chamada e documentar; requisito atendido.

**Resultado da Fase A:** Escopo do documento 100% definido (e implementado onde for necessário).

---

## Fase B – Configuração e ambiente

### B.1 Arquivo `.ini` em produção

- Copiar `.ini.example` para `.ini` na máquina/ambiente onde o RPA vai rodar.
- Preencher **todos** os parâmetros com valores reais:
  - `[holmes]`: `id_fluxo`, `token`, `id_usuario`, `id_pendencia_tarefa`, `acao_sucesso`, `acao_erro`, `nome_processo`.
  - `[webhook]`: `base_url`, `token`.
  - `[nbs]`: `caminho_sistema`, `usuario`, `senha`, `servidor`, `planilha_tipos_pagamento`, `planilha_empresas`.
  - `[email.enviar]` e `[email.destinatarios]`: sucesso, erro, **fiscal** (para TE07).
- Para **produção**, deixar:
  - `parar_apos_observacoes = false`
  - `data_corte` e `data_corte_bloquear` conforme regra de negócio (se usarem).

### B.2 Acesso e dependências na máquina do cliente

- **VPN** ativa para o ambiente onde o NBS se conecta.
- **NBS** instalado e acessível (mesmo caminho/config que no `.ini`).
- **Planilha** `TIPOS_DE_PAGAMENTO.xlsx` (ou .xls) acessível no caminho configurado em `planilha_tipos_pagamento`.
- **Planilha** `Empresas_NBS.xlsx` (se usada para fallback Empresa/Filial) acessível.
- **Resolução e escala do Windows** iguais às usadas quando você capturou as imagens de fallback (para reconhecimento de imagem).

### B.3 Imagens de fallback (recomendado)

- Confirmar que existem em `imagens/` (com os nomes usados no código):
  - `nbs_incluir_entrada_diversa.PNG` (ou .png)
  - `nbs_aba_contabilizacao.PNG`, `nbs_botao_incluir.PNG`, `nbs_aba_faturamento.PNG`, `nbs_botao_gerar.PNG`, `nbs_botao_confirmar.PNG`
  - `nbs_cnpj_nao_encontrado.PNG` (para TE01)
- O código já trata `.PNG`; se você tiver só `.png`, o `procurar_imagem` pode aceitar – senão, padronizar extensão/nome.

**Resultado da Fase B:** Ambiente pronto para rodar e testar o fluxo completo.

---

## Fase C – Teste e validação no NBS

Fazer isso **na máquina do cliente (ou com VPN + NBS)**.

### C.1 Teste com parada antecipada (opcional)

- Deixar `parar_apos_observacoes = true`.
- Rodar o RPA com **um** processo de teste.
- Validar até a tela de Total Nota e Observações: login, Empresa/Filial, Incluir Entrada (Só Diversa), CPF, Nº NF, Série, Emissão, Entrada, desmarcar livro fiscal, Total Nota, Observações.
- Se algo falhar (campo errado, ordem de Tab): anotar e ajustar no código (ordem de Tab, cliques ou imagens).

### C.2 Teste fluxo completo

- Colocar `parar_apos_observacoes = false`.
- Rodar com um ou poucos processos de teste.
- Validar:
  - Contabilização (Conta UAB, Centro de Custo).
  - Faturamento (Parcelas = 1, Gerar, Boleto, Outras Despesas, Vencimento).
  - Confirmar, pop-up (num_controle), Ficha de Controle → Cancelar.
  - Tarefa no Holmes encaminhada (sucesso).
  - E-mail de sucesso com log.
- Se a interface “escapar” (pywinauto não achar elemento): usar as imagens de fallback; se ainda faltar alguma, capturar novo print e incluir em `imagens/` com o nome esperado pelo código.

### C.3 Testes de exceção (recomendado)

- **TE01:** Processo com CPF inexistente no NBS → deve pendenciar com motivo “CNPJ não encontrado”.
- **TE07:** Processo com Tipo de Pagamento que não está na planilha → não deve lançar no NBS; deve pendenciar e enviar e-mail para Fiscal (configurar `fiscal` no `.ini`).

**Resultado da Fase C:** Fluxo e exceções validados no ambiente real.

---

## Fase D – Ajustes e entrega

### D.1 Ajustes após o teste

- Ajustar ordem de Tab, tempo de espera (`sleep`) ou seletores em `modulos/nbs/solicitacao_pagamento.py` se algo falhar no teste.
- Incluir ou trocar imagens em `imagens/` se alguma tela do NBS for diferente.
- Atualizar `PANORAMA_VS_REQUISITOS.md` e, se quiser, `README.md` com “Validado em produção em [data]”.

### D.2 Entrega considerada 100%

- **Escopo:** Itens A.1 e A.2 definidos (e implementados ou formalmente descartados).
- **Config:** `.ini` e ambiente (B) ok.
- **Validação:** Pelo menos um fluxo completo (C.2) e, se possível, TE01/TE07 (C.3) testados com sucesso.
- **Documentação:** README e panorama atualizados; se o cliente tiver checklist de entrega, preencher com base neste guia.

### D.3 O que entregar ao cliente / gestor

- Código (repositório ou pacote).
- `.ini.example` preenchido como modelo (dados sensíveis com placeholder).
- Documentação: README, CONFIGURACAO.md, INI_GUIA.md, PANORAMA_VS_REQUISITOS.md.
- Breve instrução de execução (como rodar o RPA, onde colocar `.ini`, necessidade de VPN e NBS).

---

## Resumo em checklist

| # | Ação | Feito? |
|---|------|--------|
| A.1 | Definir com negócio: leitura de Nota anexada (lista de dados ou “não usar”) | ☐ |
| A.2 | Definir com negócio: banco de controle (modelo + conexão ou “não usar”) | ☐ |
| B.1 | `.ini` de produção preenchido; `parar_apos_observacoes = false` | ☐ |
| B.2 | VPN, NBS, planilhas e resolução conferidos na máquina do cliente | ☐ |
| B.3 | Imagens de fallback em `imagens/` conferidas | ☐ |
| C.1 | (Opcional) Teste com `parar_apos_observacoes = true` | ☐ |
| C.2 | Teste fluxo completo no NBS (até Ficha Controle → Cancelar) | ☐ |
| C.3 | (Recomendado) Teste TE01 e TE07 | ☐ |
| D.1 | Ajustes de Tab/seletores/imagens após teste | ☐ |
| D.2 | Panorama/README atualizados | ☐ |
| D.3 | Pacote e documentação entregues ao cliente/gestor | ☐ |

Quando todos os itens aplicáveis estiverem feitos e o fluxo validado no ambiente real, o projeto pode ser considerado **100% entregue** em relação ao documento de requisitos.
