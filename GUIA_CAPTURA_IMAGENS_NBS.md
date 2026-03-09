# Guia - Captura de Imagens para Fallback NBS (Solicitação Pagamento Geral)

Se o pywinauto falhar (interface alterada), o RPA pode usar reconhecimento de imagem. Para isso, é preciso capturar prints dos elementos principais.

## Como capturar

1. **Resolução**: Use a mesma resolução em que o RPA vai rodar (ex.: 1920x1080).
2. **Escala do Windows**: 100% (evite 125%, 150%).
3. **Recorte**: Capture apenas o elemento (botão, aba, label) – imagens menores reconhecem melhor.
4. **Formato**: PNG, salvo em `imagens/`.

## Elementos a capturar (ordem do fluxo)

| Arquivo sugerido | O que capturar | Tela |
|------------------|----------------|------|
| `nbs_incluir_entrada_diversa.PNG` | **Segundo ícone** da toolbar (documento com +) na janela Entradas | Prioridade – usado para abrir o formulário |
| `nbs_admin_menu.PNG` | Item "Admin" ou "NBS Fiscal" no menu | Após login |
| `nbs_aba_contabilizacao.PNG` | Aba "Contabilização" | Janela Entrada |
| `nbs_botao_incluir.PNG` | Botão "+" ou "Incluir" na Contabilização | Aba Contabilização |
| `nbs_aba_faturamento.PNG` | Aba "Faturamento" | Janela Entrada |
| `nbs_botao_gerar.PNG` | Botão "Gerar" | Aba Faturamento |
| `nbs_botao_confirmar.PNG` | Botão "Confirmar" | Janela Entrada |
| `nbs_cnpj_nao_encontrado.PNG` | Mensagem de erro "não encontrado" / "não cadastrado" | Pop-up TE01 |
| `nbs_popup_ok_sucesso.PNG` | Botão "OK" do pop-up de sucesso | Após confirmar |
| `nbs_ficha_controle_cancelar.PNG` | Botão "Cancelar" da Ficha de Controle | Ficha de Controle |

## Dicas

- **Botões**: Recorte o botão inteiro, com um pouco de margem.
- **Abas**: Recorte a aba com o texto visível.
- **Pop-ups**: Recorte a mensagem ou o botão OK.
- **Evite**: Elementos que mudam (datas, valores, números de processo).

## Após capturar

1. Coloque os arquivos em `imagens/` com os nomes exatos da tabela acima.
2. O fallback já está integrado: o fluxo tenta pywinauto primeiro; se não encontrar o elemento, usa a imagem.
3. Você pode adicionar as imagens aos poucos – cada uma passa a funcionar assim que o arquivo existir.
