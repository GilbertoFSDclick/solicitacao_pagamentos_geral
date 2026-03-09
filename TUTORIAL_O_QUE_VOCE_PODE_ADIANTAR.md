# Tutorial: O que você pode adiantar (passo a passo)

Tudo que **depende só de você** antes de ter VPN, máquina do cliente ou respostas do negócio. Siga na ordem.

---

## Passo 1 – Deixar o projeto pronto para rodar

### 1.1 Criar seu arquivo `.ini` local

1. Na pasta raiz do projeto, copie o arquivo de exemplo:
   - **De:** `.ini.example`  
   - **Para:** `.ini`

2. Abra o `.ini` no editor.

3. Preencha o que você **já tiver** hoje; o que não tiver, deixe como está ou use placeholders claros (ex.: `ID_A_PREENCHER`).

**O que costuma dar para preencher mesmo sem VPN:**

| Seção | Chave | O que colocar |
|-------|--------|----------------|
| `[email.enviar]` | user, password, host | Seu e-mail de envio (ex.: Outlook) |
| `[email.destinatarios]` | sucesso, erro, fiscal | E-mails que receberão os relatórios |
| `[bot]` | nome | Nome do bot (ex.: "Lançamento Sol. Pagamento Geral - NBS") |
| `[holmes]` | host | Normalmente `https://app-api.holmesdoc.io` |
| `[webhook]` | base_url, max_tentativas | base_url do serviço de RPA; max_tentativas = 3 |
| `[nbs]` | caminho_sistema, titulo_sistema | Caminho do NBS e título da janela (quando tiver) |
| `[logger]` | flag_persistencia, dias_persistencia | true e 14 |

**O que normalmente só o cliente/infra tem (deixe placeholder):**

- `[holmes]`: id_fluxo, token, id_usuario, id_pendencia_tarefa, acao_sucesso, acao_erro, nome_processo
- `[webhook]`: token
- `[nbs]`: usuario, senha, servidor, planilha_empresas, planilha_tipos_pagamento

4. No `[nbs]`, deixe **`parar_apos_observacoes = true`** até fazer o primeiro teste no NBS (depois mude para `false`).

5. Salve o `.ini`.  
   (O `.ini` não vai para o Git; está no `.gitignore`.)

---

## Passo 2 – Conferir imagens de fallback

O RPA usa imagens da pasta `imagens/` quando não acha os elementos pela tela (pywinauto).

### 2.1 Nomes que o código espera

O código usa estes caminhos (em maiúsculas `.PNG`):

- `imagens/nbs_incluir_entrada_diversa.PNG`
- `imagens/nbs_aba_contabilizacao.PNG`
- `imagens/nbs_botao_incluir.PNG`
- `imagens/nbs_aba_faturamento.PNG`
- `imagens/nbs_botao_gerar.PNG`
- `imagens/nbs_botao_confirmar.PNG`
- `imagens/nbs_cnpj_nao_encontrado.PNG`

No Windows, `.png` e `.PNG` costumam ser tratados iguais. Em servidor Linux, o nome deve ser exatamente o que o código usa.

### 2.2 O que fazer agora

1. Abra a pasta **`imagens`** do projeto.

2. Para cada arquivo que você já tiver (ex.: `nbs_incluir_entrada_diversa.png`):
   - Se o nome estiver diferente do da lista acima, **renomeie** para o nome esperado (ex.: `nbs_incluir_entrada_diversa.PNG`).
   - Se faltar alguma imagem da lista, anote num bloco de notas: “Falta: nbs_aba_contabilizacao.PNG, …” para capturar quando estiver no NBS.

3. As que ainda não existem não impedem o RPA de rodar: o código tenta pela tela primeiro e só usa a imagem se existir. Você pode ir adicionando depois.

---

## Passo 3 – Testar se o ambiente Python está ok

Objetivo: garantir que o projeto abre e que o `.ini` é lido, **sem** precisar do NBS ou da VPN.

### 3.1 Abrir o terminal na pasta do projeto

- No VS Code/Cursor: Terminal → “New Terminal” (ou `Ctrl+'`).
- Confirme que o caminho é a pasta do projeto (onde está o `main.py`).

### 3.2 Ativar o ambiente virtual (se existir)

Se existir a pasta `venv`:

- **Windows (PowerShell):**  
  `.\venv\Scripts\Activate.ps1`
- **Windows (CMD):**  
  `.\venv\Scripts\activate.bat`

O prompt deve aparecer algo como `(venv)` na frente.

Se não existir `venv`, crie e instale as dependências:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 3.3 Rodar um “teste seco” (sem processar processos de verdade)

No terminal, execute:

```powershell
python -c "
import bot
import operacoes
from operacoes.tratar_tarefa import MAPEAMENTO_CAMPOS
print('OK: imports e .ini carregados')
print('Campos mapeados do Holmes:', list(MAPEAMENTO_CAMPOS.keys())[:5], '...')
"
```

- Se aparecer **“OK: imports e .ini carregados”** e a lista de campos, o ambiente e o `.ini` estão sendo lidos.
- Se der erro de “módulo não encontrado”, instale as dependências: `pip install -r requirements.txt`.
- Se der erro de “opção não encontrada” ou “config”, confira o `.ini` (seções e nomes das chaves iguais ao `.ini.example`).

---

## Passo 4 – Preparar a lista de perguntas para o negócio

Isso adianta a Fase A do “Entrega 100%” e evita travamento depois.

### 4.1 Sobre “Leitura de Nota anexada”

Mande um e-mail ou mensagem para o gestor/negócio com algo assim:

- “No documento do RPA está: ‘Leitura de Nota anexada ao processo – lista de dados a definir’. Precisamos saber: (1) Quais dados exatamente devemos extrair da nota? (2) É um anexo específico do processo no Holmes (qual nome/tipo)? (3) Se não for usar essa leitura, podemos seguir sem ela?”

Anote a resposta para implementar (ou registrar que não será usado).

### 4.2 Sobre “Banco de controle”

Outra mensagem possível:

- “O documento fala em ‘Inclusão em banco de dados de controle – dados a definir’. Precisamos: (1) Confirmação se haverá banco de controle. (2) Se sim: qual tabela, quais colunas e quem fornece a conexão (string, ambiente)?”

Anote a resposta para implementar o insert ou desativar a função.

---

## Passo 5 – Preparar o “kit” para quando tiver VPN / NBS

Assim, no dia em que tiver acesso, você só segue um checklist.

### 5.1 Criar um checklist de primeiro teste

Crie um arquivo de texto ou use o bloco de notas com algo assim (pode colar no projeto como `CHECKLIST_PRIMEIRO_TESTE.txt`):

```
[ ] VPN conectada
[ ] NBS instalado no caminho do .ini
[ ] .ini com usuario/senha/servidor NBS preenchidos
[ ] planilha_tipos_pagamento e planilha_empresas acessíveis (caminho do .ini)
[ ] parar_apos_observacoes = true (primeiro teste)
[ ] Rodar: python main.py
[ ] Verificar até tela Total Nota e Observações
[ ] Se der erro em algum campo: anotar qual tela e qual campo
[ ] Depois: parar_apos_observacoes = false e testar fluxo completo
```

### 5.2 Deixar documentação à mão

Na pasta do projeto você já tem:

- **ENTREGA_100_PORCENTO.md** – o que falta para 100%
- **PANORAMA_VS_REQUISITOS.md** – estado vs. documento
- **CONFIGURACAO.md** / **INI_GUIA.md** – se existirem, use para conferir o `.ini`

Abra esses arquivos quando for fazer o primeiro teste na VPN.

---

## Passo 6 – (Opcional) Documentar onde está cada coisa

Para você ou outra pessoa não se perder depois:

1. Onde fica o **código do fluxo NBS**  
   - `modulos/nbs/solicitacao_pagamento.py` (função `processar_entrada_solicitacao_pagamento`).

2. Onde fica o **mapeamento dos campos do Holmes**  
   - `operacoes/tratar_tarefa.py` (constante `MAPEAMENTO_CAMPOS` e função `tratar_tarefa_aberta`).

3. Onde fica a **planilha de tipos de pagamento**  
   - Leitura em `modulos/nbs/solicitacao_pagamento.py`, função `obter_codigo_uab_contabil`; caminho no `.ini`: `[nbs] planilha_tipos_pagamento`.

4. Onde ativar/desativar **parada após Observações**  
   - No `.ini`: `[nbs] parar_apos_observacoes = true` ou `false`.

Isso você pode colocar em um “README interno” ou no próprio **README.md** do projeto em 4 linhas.

---

## Resumo: o que você adiantou

| Passo | O que fez |
|-------|-----------|
| 1 | `.ini` criado e preenchido no que der; `parar_apos_observacoes = true` |
| 2 | Imagens conferidas/renomeadas; lista do que falta anotada |
| 3 | Ambiente Python e leitura do `.ini` testados |
| 4 | Perguntas para o negócio (nota anexada + banco) enviadas e respostas anotadas |
| 5 | Checklist de primeiro teste e doc à mão para o dia da VPN |
| 6 | (Opcional) Anotado onde fica fluxo NBS, Holmes, planilha e parada |

Quando tiver VPN e NBS, use o **CHECKLIST_PRIMEIRO_TESTE** e o **ENTREGA_100_PORCENTO.md** para fechar o que falta e considerar o projeto 100% entregue.
