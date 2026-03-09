# Guia de Configuração do .ini

## Campos obrigatórios por seção

### [email.enviar]
| Campo | Descrição |
|-------|------------|
| user | E-mail para envio (SMTP) |
| password | Senha do e-mail |
| host | Servidor SMTP (ex: smtp-mail.outlook.com) |

### [email.destinatarios]
| Campo | Descrição |
|-------|------------|
| sucesso | E-mails que recebem relatório de sucesso (separados por vírgula) |
| erro | E-mails que recebem alerta de erro |
| fiscal | E-mails para TE07 (Tipo de Pagamento não parametrizado) |

### [holmes]
| Campo | Descrição | Onde obter |
|-------|------------|------------|
| host | URL da API Holmes | Padrão: https://app-api.holmesdoc.io |
| token | Token de autenticação | Holmes > Configurações / API |
| id_fluxo | ID do fluxo "Solicitação de Pagamentos Geral" | Holmes > Editar fluxo > URL ou propriedades |
| nome_processo | Nome exato do processo | Ex: "Tarefas - Solicitação de Pagamentos Geral" |
| nome_atividade | Nome da tarefa que a RPA assume | Ex: "Departamento de Multas" ou "Lançamento RPA" |
| id_usuario | ID do usuário RPA no Holmes | Holmes > Usuários |
| id_pendencia_tarefa | ID do campo de pendência | Holmes > Propriedades do fluxo |
| acao_sucesso | ID da ação de sucesso (Finalizar) | Holmes > Ações da tarefa |
| acao_erro | ID da ação de erro (Pendenciar) | Holmes > Ações da tarefa |

### [webhook]
| Campo | Descrição |
|-------|------------|
| base_url | URL do serviço de webhook |
| token | Token de autenticação do webhook |
| max_tentativas | Número de retentativas antes de pendenciar (ex: 3) |

### [nbs]
| Campo | Descrição |
|-------|------------|
| caminho_sistema | Caminho do executável do NBS (ex: C:\NBS\SisFin.exe) |
| titulo_sistema | Título da janela do NBS |
| planilha_empresas | Caminho da planilha Empresas_NBS (de/para CNPJ → Código Empresa/Filial) |
| planilha_tipos_pagamento | Caminho da planilha TIPOS_DE_PAGAMENTO (Col A → Col D) |
| usuario | Usuário de login no NBS |
| senha | Senha do NBS |
| servidor | Servidor NBS (ex: uab) |

## Para teste local

- `planilha_tipos_pagamento = ..\TIPOS_DE_PAGAMENTO.xlsx` (caminho relativo à pasta do projeto)
- Mantenha os demais valores conforme ambiente do cliente
