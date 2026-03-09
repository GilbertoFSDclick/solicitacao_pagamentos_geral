# Guia: Deploy Seguro no GitHub e na VM do Cliente

Este guia orienta o envio do projeto para o GitHub e a instalação na máquina virtual do cliente de forma **segura**, sem expor credenciais ou dados sensíveis.

---

## 1. Antes de subir no GitHub

### 1.1 O que NUNCA deve ir para o repositório

| Arquivo/Conteúdo | Motivo |
|------------------|--------|
| **`.ini`** | Contém senhas, tokens, e-mails e credenciais |
| **`*.log`** | Podem conter dados de processos e erros |
| **`controle_rpa.db`** (ou `*.db`) | Banco com dados do cliente |
| **Planilhas** (`*.xlsx`, `*.xls`) | Podem ter dados sensíveis |
| **Credenciais em código** | Nada de senhas/tokens hardcoded |

### 1.2 Verificar o .gitignore

O projeto já possui `.gitignore` com:

- `.ini` ✅
- `*.log` ✅
- `**/__pycache__` ✅
- `*.xlsx` ✅

**Recomendação:** Adicione `*.db` para garantir que o banco SQLite não seja commitado:

```
# banco de dados (dados do cliente)
*.db
```

### 1.3 Conferir se o .ini está fora do Git

Antes do primeiro commit:

```powershell
git status
```

Se `.ini` aparecer como "untracked" ou "modified", **não** faça `git add .ini`. O `.gitignore` deve impedir isso, mas confira.

---

## 2. Criar o repositório no GitHub

### 2.1 Repositório privado (recomendado)

1. Acesse [github.com](https://github.com) e faça login
2. **New repository**
3. Nome sugerido: `rpa-solicitacao-pagamento-nbs` (ou similar)
4. **Marque "Private"** – evita que o código fique público
5. Não marque "Initialize with README" (o projeto já existe)
6. Clique em **Create repository**

### 2.2 Conectar e enviar o projeto

Na pasta do projeto (onde está o `main.py`):

```powershell
# Se ainda não inicializou o Git
git init

# Adicionar o remote (substitua SEU_USUARIO e NOME_REPO pelos seus)
git remote add origin https://github.com/SEU_USUARIO/NOME_REPO.git

# Ver o que será commitado (confirme que .ini NÃO aparece)
git status

# Adicionar arquivos (o .gitignore exclui .ini, *.log, etc.)
git add .

# Primeiro commit
git commit -m "Projeto RPA Lançamento Solicitação Pagamento Geral - NBS"

# Enviar para o GitHub (branch main)
git branch -M main
git push -u origin main
```

### 2.3 Autenticação no GitHub

- **HTTPS:** O GitHub pedirá usuário e senha. Use um **Personal Access Token (PAT)** em vez da senha da conta.
- **SSH:** Se tiver chave SSH configurada, use: `git@github.com:SEU_USUARIO/NOME_REPO.git`

Para criar um PAT: GitHub → Settings → Developer settings → Personal access tokens → Generate new token.

---

## 3. Baixar na VM do cliente

### 3.1 Pré-requisitos na VM

- **Python 3.10+** instalado
- **Git** instalado (ou download do ZIP do repositório)
- **VPN** ativa (para acessar NBS e planilhas)
- **NBS** instalado e configurado
- Acesso à pasta de planilhas (`\\10.100.9.7\rpa\...` ou equivalente)

### 3.2 Clonar o repositório

```powershell
# Navegar até a pasta onde quer o projeto (ex: C:\RPA)
cd C:\RPA

# Clonar (substitua pela URL do seu repositório)
git clone https://github.com/SEU_USUARIO/NOME_REPO.git

cd NOME_REPO
```

Se o repositório for privado, será pedido login/token.

### 3.3 Criar o arquivo .ini (sem commitá-lo)

O `.ini` **não** vem no repositório. Crie na VM:

```powershell
# Copiar o exemplo
copy .ini.example .ini

# Editar com os valores reais (Notepad, VS Code, etc.)
notepad .ini
```

Preencha com os dados do ambiente do cliente:

- `[email.enviar]` – usuário e senha do e-mail
- `[holmes]` – token, id_fluxo, id_usuario, etc.
- `[webhook]` – token
- `[nbs]` – usuario, senha, caminhos das planilhas (ex: `\\10.100.9.7\rpa\...`)

### 3.4 Instalar dependências

```powershell
pip install -r requirements.txt
```

Se usar `sqlalchemy` para o banco de controle:

```powershell
pip install sqlalchemy
```

### 3.5 Criar a tabela do banco (se usar)

```powershell
python scripts/criar_tabela_banco_controle.py
```

### 3.6 Verificar imagens de fallback

Confirme que a pasta `imagens/` contém os arquivos esperados (ex.: `nbs_incluir_entrada_diversa.png`). Eles devem estar no repositório.

---

## 4. Checklist de segurança

| Item | Verificado? |
|------|-------------|
| `.ini` no .gitignore | ☐ |
| `.ini` não foi commitado (`git status` não mostra) | ☐ |
| Repositório privado no GitHub | ☐ |
| Nenhuma senha/token no código-fonte | ☐ |
| `*.db` no .gitignore | ☐ |
| `.ini` criado manualmente na VM com dados do cliente | ☐ |

---

## 5. Atualizações futuras

Quando fizer alterações no código:

```powershell
git add .
git commit -m "Descrição da alteração"
git push
```

Na VM do cliente:

```powershell
git pull
```

O `.ini` da VM **não** será sobrescrito pelo `git pull`, pois não está versionado.

---

## 6. Alternativa: enviar .ini por canal seguro

Se preferir não digitar o `.ini` na VM:

1. Crie o `.ini` na sua máquina com os dados do cliente
2. Envie por canal seguro (e-mail criptografado, compartilhamento seguro, etc.)
3. Coloque o arquivo na pasta do projeto na VM
4. **Nunca** faça commit desse `.ini`

---

## 7. Resumo

1. **GitHub:** Repo privado, `.gitignore` correto, sem `.ini` nem `*.db`
2. **VM:** Clone do repo, criar `.ini` localmente, instalar dependências
3. **Segurança:** Credenciais apenas no `.ini` local, nunca no repositório
