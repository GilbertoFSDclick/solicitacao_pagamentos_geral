#std
import atexit
from typing import Literal, NamedTuple, Optional, Self
import os
import xml.etree.ElementTree as ET
from pywinauto.controls.hwndwrapper import HwndWrapper
from decimal import Decimal, ROUND_HALF_UP
import holidays
import subprocess

# interno
import bot
import modulos
from bot.estruturas import Coordenada
from time import sleep
from modulos.interface import Elemento
from datetime import datetime, timedelta

#externo
import pandas as pd

class ErroNBS (Exception):
    """Erro próprio para o sistema NBS"""
    def __init__ (self, mensagem: str) -> None:
        self.mensagem = mensagem
        super().__init__(self.mensagem)

class RetornoStatus(NamedTuple):
    """Classe com padrão de retorno
    - `SUCESSO` indica se o retorno foi bem-sucedido
    - `MENSAGEM` passa detalhes em caso de mal-sucedido"""
    SUCESSO: bool
    MENSAGEM: str = ""

class Sistema:
    """Sistema NBS"""

    _atexit_encerrar_sistema = False
    """Estado atual do registro da função no `atexit` que encerra o sistema ao fim da execução"""

    def __init__(self, usuario, senha, servidor) -> Self:
        """Cria uma instância do sistema NBS"""
        self.usuario = usuario
        self.senha = senha
        self.servidor = servidor

        # Obter variáveis de acesso ao sistema
        self.caminho_sistema, self.titulo_sistema = bot.configfile.obter_opcoes("nbs", ["caminho_sistema", "titulo_sistema"])

        Sistema.encerrar()

        # Garante que o hook de encerramento é registrado apenas uma vez
        if not Sistema._atexit_encerrar_sistema:
            atexit.register(Sistema.encerrar)
            Sistema._atexit_encerrar_sistema = True

    def inicializar(self) -> Self:
        """Executa o NBS"""
       
        try:
            # Verifica caminho do sistema
            caminho_valido =  bot.windows.afirmar_arquivo( self.caminho_sistema )
            assert caminho_valido, "O caminho especificado para o NBS é inválido"

            # Executa o sistema
            bot.logger.informar('Executando NBS...')
            os.startfile( self.caminho_sistema )

            # Aguarda um tempo para sistema carregar
            sistema_abriu = bot.util.aguardar_condicao(lambda: self.titulo_sistema
                                              in bot.windows.Janela.titulos_janelas(), 30)
            if not sistema_abriu:
                raise Exception("Janela do sistema não abriu no tempo esperado")
            
            # Foca na janela do sistema
            janela = bot.estruturas.Janela( self.titulo_sistema ).focar()

            # Faz login com as credenciais
            bot.teclado.digitar_teclado(self.usuario, .05)
            bot.teclado.apertar_tecla('tab')
            bot.teclado.digitar_teclado(self.senha, .05); sleep(1)
            bot.teclado.apertar_tecla('tab')
            bot.teclado.digitar_teclado(self.servidor, .05); sleep(1)
            bot.teclado.apertar_tecla('enter')

        except Exception as erro:
            raise ErroNBS(f"Erro ao inicializar o NBS: { erro }")
        
        bot.logger.informar('NBS executado com sucesso!')
        return self
    
    @classmethod
    def encerrar(cls) -> None:
        """Encerra o sistema NBS"""
        processos = bot.configfile.obter_opcao_ou("nbs", "processos")

        comando = f"""
            Get-Process -Name {processos} -IncludeUserName |
                Where-Object {{ $_.UserName -eq "$env:USERDOMAIN\\$env:USERNAME" }} |
                Stop-Process -Force
        """
        subprocess.run(["powershell", "-Command", comando], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def valida_credito_ctb(torre: str, conta_contabil: str) -> tuple[str, str] | None:
    """Verifica se `conta_contabil` está presente na relação de contas que creditam PIS/COFINS
    e retorna os valores das colunas 'Serviço' e 'Produto' correspondentes como uma tupla de strings.
    Caso a conta não seja encontrada ou ocorra um erro, retorna None.
    """
    bot.logger.informar(f'Verificando classificação contábil {conta_contabil} no de/para')
    try:
        planilha_contas_ctb = bot.configfile.obter_opcao_ou('nbs', 'planilha_contas_ctb')
        df = pd.read_excel(planilha_contas_ctb, sheet_name=torre, skiprows=1)

        df.columns = ['Classificação Contabil', 'Descricao', 'Servico', 'Produto']
        
        # Normalização
        df["Classificação Contabil"] = df['Classificação Contabil'].astype(str).apply(bot.util.normalizar)
        conta_contabil = bot.util.normalizar(str(conta_contabil))
        
        # Verifica se a conta_contabil está na coluna Classificação Contabil
        mask = df['Classificação Contabil'] == conta_contabil
        if mask.any():
            # Linha correspondente
            linha = df[mask]

            # Extrai os valores das colunas 'Servico' e 'Produto'
            servico = str(linha['Servico'].iloc[0]) # Coluna Servico
            produto = str(linha['Produto'].iloc[0]) # Coluna Produto

            bot.logger.informar(f'Conta { conta_contabil } identificada na lista.')
            return (servico, produto)
        else:
            bot.logger.informar(f'Conta { conta_contabil } não identificada na lista.')
            return None
    except Exception as erro:
        bot.logger.erro(f'Erro ao validar conta contábil na planilha: { erro }')
        return None

def consultar_de_para_empresa(
            cnpj_filial: str,
            aba: str,
            coluna: Literal['Descrição Patio NBS', 'Cod_empresa',
                            "Nbs_Filial", "CNPJ", "Razão Social"]) -> Optional[str]:
        """Retorna a informação da `coluna` de acordo com `cnpj_filial` (CNPJ) na `aba` (sheet) especificada"""

        planilha = bot.configfile.obter_opcao_ou('nbs', 'planilha_empresas')
        if not bot.windows.afirmar_arquivo(planilha):
            raise Exception("Não foi possível acessar a planilha. O caminho para o arquivo é inválido.")

        bot.logger.informar(f'Obtendo { coluna } da filial de cnpj { cnpj_filial }')

        df = pd.read_excel(planilha, sheet_name=aba, dtype={"CNPJ": str})
        df.columns = df.columns.str.strip()

        empresa = df[df['CNPJ'].astype(str).str.strip().str.lower().str.replace(' ', '') == str(cnpj_filial).strip().lower().replace(' ', '')]

        return str(empresa[coluna].values[0]).strip() if not empresa.empty else None

def processar_entrada(campos_obrigatorios:list, propriedades_processo) -> RetornoStatus | bool:
    """Realiza a entrada de nota no NBS
    - `campos_obrigatorios` todos os campos usados durante o processamento"""

    nomeFilial, numeroNf, valorNf, cnpj_filial, centroCusto, temParcelamento, classificacaoContabil, torre, dataVencimento, quantidadeParcelas, cnpjEmissor, protocolo, tipo_doc = campos_obrigatorios
    
    # Opções de Natureza de Crédito
    de_para_natureza_credito: dict[str, str] = {
        '1': 'Aquisicao de bens para revenda',
        '2': 'Aquisicao de bens utilizados como insumo',
        '3': 'Aquisicao de servicos utilizados como insumo',
        '4': 'Energia eletrica e termica',
        '5': 'Alugueis de predios',
        '6': 'Alugueis de maquinas e equipamentos',
        '7': 'Armazenagem de mercadoria e frete na operacao de venda',
        '8': 'Contraprestacoes de arrendamento mercantil'
    }

    bot.logger.informar(f"CNPJ Filial: { cnpj_filial }")

    # Início do processamento
    try:
        bot.logger.informar(f"Iniciando entrada da nota '{ numeroNf }' na filial '{ nomeFilial }'")
        bot.estruturas.Janela('NBS Sistema Financeiro').focar()

        # obter código da Empresa NBS
        cod_empresa: str = consultar_de_para_empresa(cnpj_filial, torre, 'Nbs_empresa', )
        assert cod_empresa, "Código da empresa não encontrado"

        # obter código da Filial NBS
        cod_filial: str = consultar_de_para_empresa(cnpj_filial, torre, 'Nbs_Filial')
        assert cod_filial, "Código da filial não encontrado"
        
        bot.logger.informar(f"Digitando codigo da empresa: { cod_empresa } e filial {cod_filial}")
        # Digita número da empresa e confirma em 'OK'
        bot.teclado.digitar_teclado(cod_empresa)
        bot.teclado.apertar_tecla('tab')
        bot.teclado.digitar_teclado(cod_filial)
        bot.teclado.apertar_tecla('tab')
        bot.teclado.apertar_tecla('enter')

        #Aguardar Janela
        bot.util.aguardar_condicao(lambda: "Sistema Financeiro - SISFIN" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Sistema Financeiro - SISFIN")
        bot.logger.informar("Selecionando Contas a Pagar")
        #Selecionar Contas a Pagar
        bot.teclado.apertar_tecla('alt')
        bot.teclado.digitar_teclado('P')
        bot.teclado.digitar_teclado('C')
        bot.teclado.apertar_tecla('enter')

        #Aguardar Janela
        bot.util.aguardar_condicao(lambda: "Contas a Pagar" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Contas a Pagar")

        #Clicar em 'Nota Fiscal de Compra'
        imagem_nota_fiscal = bot.imagem.procurar_imagem("imagens/nota-fiscal-compra.PNG", confianca=0.8, segundos=5)
        if imagem_nota_fiscal: 
            bot.mouse.clicar_mouse(coordenada=imagem_nota_fiscal)
            bot.logger.informar("Entrando Nota FIscal de Compra")
        else:
            janela.fechar() 
            raise Exception("Erro ao abrir 'Nota Fiscal de Compra")
        
        #Aguardar Janela
        bot.util.aguardar_condicao(lambda: "Entradas" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Entradas")
        
        #Clicar em 'NFE"
        imagem_nfe = bot.imagem.procurar_imagem("imagens/nfe.PNG", confianca=0.8, segundos=5)
        if imagem_nfe: 
            bot.mouse.clicar_mouse(coordenada=imagem_nfe)
            bot.logger.informar("Clicando em 'NFE' ")

        else:
            janela.fechar() 
            raise Exception("Erro ao abrir 'Nota Fiscal de Compra")
        
        #Clicar e preencher Nota Fiscal
        bot.util.aguardar_condicao(lambda: "Arquivos NFe: Interface de compra" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Arquivos NFe: Interface de compra")
        elemento = janela.elementos(class_name="TOvcPictureField", top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.digitar_teclado(numeroNf); sleep(0.5)

        #Desmarcar data de emissão
        bot.teclado.apertar_tecla('tab')
        bot.teclado.apertar_tecla('space')

        #Clicar em Monitor NF-e
        imagem_monitor_nfe = bot.imagem.procurar_imagem("imagens/monitor-nfe.PNG", confianca=0.8, segundos=5)
        if imagem_monitor_nfe:
            bot.mouse.clicar_mouse(coordenada=imagem_monitor_nfe)
            bot.logger.informar("Clicando em 'Monitor NFE' ")     
        else:
            janela.fechar()
            raise Exception("Erro ao abrir Monitor NFE") 
        
        bot.util.aguardar_condicao(lambda: "Monitor Notas Eletrônicas - NFe Compra Diversa" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Monitor Notas Eletrônicas - NFe Compra Diversa")
        sleep(2)

        #Clicar na data Emissão "até"
        elemento = janela.elementos(class_name="TDateTimePicker", top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        
        #Clicar na data de emissão
        # bot.teclado.apertar_tecla('tab',quantidade=6,delay=0.3)
        # bot.teclado.apertar_tecla('space')
        # bot.teclado.apertar_tecla()
        bot.teclado.atalho_teclado(['shift', 'tab'])
        bot.teclado.apertar_tecla('space')


        #Clicar e preencher Nota Fiscal
        elemento = janela.elementos(class_name="TOvcPictureField", top_level_only=False)[1]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.digitar_teclado(numeroNf); sleep(0.5)

        vIpi_tag, vFrete_tag, vDesc_tag, serie, icms_st_tag, outros_tag  = valores_xml()
        #Clicar e preencher Série da Nota Fiscal
        elemento = janela.elementos(class_name="TOvcPictureField", top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.digitar_teclado(serie); sleep(0.5)

        #Clicar e preencher "Fornecedor" com o CNPJ do Emissor
        elemento = janela.elementos(class_name="TEdit", top_level_only=False)[2]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.digitar_teclado(cnpjEmissor); sleep(0.5)
        
        #Clicar em Pesquisar
        bot.teclado.apertar_tecla('f1'); sleep(4)

        #  #Clicar NÃO salavr nota
        # bot.teclado.apertar_tecla('n')

        #Clicar no Painel de notas TwwDBGrid
        elemento = janela.elementos(class_name="TwwDBGrid", top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.apertar_tecla('space')
        bot.teclado.apertar_tecla('enter')

        # Verificar Caixa de Dialogo
        while True:
            confirmacao = bot.util.aguardar_condicao(lambda: "Confirmação" in bot.windows.Janela.titulos_janelas(), 3)
            if not confirmacao:
                break 

            janela_confirmacao = bot.windows.Janela("Confirmação")
            janela_confirmacao.focar()
            bot.teclado.atalho_teclado(['ctrl', 'c'])
            texto_copiado = bot.teclado.texto_copiado()

            palavras_remover = ["Erro", "OK", "-"]
            for palavra in palavras_remover:
                texto_copiado = texto_copiado.replace(palavra, "")
            texto = " ".join(texto_copiado.split())

            if "cancelada com erro" in texto:
                bot.logger.alertar(f"Confirmação detectada: {texto}")
                bot.teclado.apertar_tecla("s")
                
            elif "Não verificada junto à SEFAZ" in texto:
                bot.logger.informar("Confirmação SEFAZ detectada, clicando em 's'")
                bot.teclado.apertar_tecla("s")
            else:
                break

        confirmacao = bot.util.aguardar_condicao(lambda: "Confirmação" in bot.windows.Janela.titulos_janelas(), 5)

        if confirmacao:
            sleep(1)
            bot.teclado.apertar_tecla('s')
        else:
            bot.teclado.apertar_tecla('enter')
            raise ErroNBS(f"Erro ao procurar a nota {numeroNf} no Painel de Notas")

        #Clicar em Pesquisar
        imagem_pesquisar = bot.imagem.procurar_imagem("imagens/pesquisar.PNG", confianca=0.8, segundos=5)
        if imagem_pesquisar: 
            bot.mouse.clicar_mouse(coordenada=imagem_pesquisar)
            bot.logger.informar("Clique em Pesquisar")     
        else:
            janela.fechar()
            raise Exception("Não foi possivel clicar em pesquisar") 

        bot.util.aguardar_condicao(lambda: "Arquivos NFe: Interface de compra" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Arquivos NFe: Interface de compra")
        

        #Clicar em Aceitar
        imagem_aceitar = bot.imagem.procurar_imagem("imagens/aceitar.PNG", confianca=0.8, segundos=5)
        if imagem_aceitar: 
            bot.mouse.clicar_mouse(coordenada=imagem_aceitar)    
            bot.logger.informar("Clique em aceitar")     
        else:
            janela.fechar()
            raise Exception("Não foi possivel clicar em pesquisar") 

        bot.util.aguardar_condicao(lambda: "Entrada Diversas / Operação: 52-Entrada Diversas" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Entrada Diversas / Operação: 52-Entrada Diversas")

        #Procurar erro de Fornecedor não cadastrado
        imagem_erro_fornecedor = bot.imagem.procurar_imagem("imagens/fornecedor_nao_cadastrado.PNG", confianca=0.8, segundos=5)
        if imagem_erro_fornecedor: 
            janela.fechar()
            raise ErroNBS(f"Falha na Leitura do XML: Fornecedor não cadastrado")      
        else: bot.logger.informar("Não encontrei mensagem de erro, seguindo...") 
        
        #Clicar em capa
        imagem_capa = bot.imagem.procurar_imagem("imagens/capa.PNG", confianca=0.8, segundos=3)
        if not imagem_capa: raise Exception("Não foi possivel clicar em 'capa'")
        bot.mouse.clicar_mouse(coordenada=imagem_capa)
        sleep(2)
        bot.logger.informar(f"PROTOCOLO {protocolo}")

        #Preencher protocolo
        bot.logger.informar("Preencher protocolo")
        elemento = janela.elementos(title='      Capa      ', top_level_only=False)[0]
        elementos: list[HwndWrapper] = elemento.children(class_name='TPanel')[-1]
        executar = Elemento(elementos)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.digitar_teclado(protocolo)

        # Regra de de crédito PIS/COFINS de acordo com planilha
        valida_credito = valida_credito_ctb(torre=torre, conta_contabil=classificacaoContabil)
        if valida_credito:
            cod_nat_servico = valida_credito[0]
            cod_nat_produto = valida_credito[1]

            janela = bot.windows.Janela()

            # Marca na opção PIS/COFINS
            botao_pis = janela.elementos(title="PIS", top_level_only=False)[0]
            assert botao_pis, "Opção Creditar PIS não identificada"
            bot.mouse.clicar_mouse(coordenada=Elemento(botao_pis).coordenada)

            sleep(2)

            # Clica na aba Natureza Créditos
            aba_natureza_creditos = janela.elementos(title="Natureza Créditos Pis/Cofins", top_level_only=False)[0]
            bot.mouse.clicar_mouse(coordenada=Elemento(aba_natureza_creditos).coordenada)

            if str(tipo_doc).lower() == 'despesas serviços':
                # Foca no campo 'Nat. BC de Crédito Serviço - Sped Pis/Cofins'
                campo = janela.elementos(title="Nat. BC de Crédito Produto - Sped Pis/Cofins", top_level_only=False)[0].children()[-1]
                bot.mouse.clicar_mouse(coordenada=Elemento(campo).coordenada)

                bot.teclado.digitar_teclado(de_para_natureza_credito[cod_nat_produto])
                bot.teclado.apertar_tecla("enter")

            elif str(tipo_doc).lower() == 'despesas produtos':
                # Foca no campo 'Nat. BC de Crédito Produto - Sped Pis/Cofins'
                campo = janela.elementos(title="Nat. BC de Crédito Produto - Sped Pis/Cofins", top_level_only=False)[0].children()[-1]
                bot.mouse.clicar_mouse(coordenada=Elemento(campo).coordenada)
                bot.teclado.digitar_teclado(de_para_natureza_credito[cod_nat_produto])
                bot.teclado.apertar_tecla("enter")
                
        #Se a data for maior que 28 as 12:00 lança como dia 1 do mês seguinte
        data = modulos.nbs.comparar_data()
        if data:
            elemento = janela.elementos(class_name='TOvcDbPictureField', top_level_only=False)[3]
            executar = Elemento(elemento)
            bot.mouse.clicar_mouse(coordenada=executar.coordenada)
            bot.teclado.apertar_tecla('right', quantidade=10)
            bot.teclado.apertar_tecla('backspace', quantidade=10)
            bot.teclado.digitar_teclado(data)

        #Clicar em Cfops
        imagem_cfops = bot.imagem.procurar_imagem("imagens/cfops.PNG", confianca=0.8, segundos=5)
        if imagem_cfops: 
            bot.mouse.clicar_mouse(coordenada=imagem_cfops)   
            bot.logger.informar("Clicando em 'CFOPS' ")  
        else:
            janela.fechar()
            raise Exception("Não foi possivel clicar em 'CFOPS' ") 

        #Clicar e preencher Código de Natureza/Grupo
        #Leitura do XML
        produtos = produtos_xml()
        #Verificando se tem frete e Ipi
        vIpi_tag, vFrete_tag, vDesc_tag, serie,icms_st_tag, outros_tag = valores_xml()
        if vIpi_tag is None: vIpi_tag = '0.00'
        if vFrete_tag is None: vFrete_tag = '0.00'
        if vDesc_tag is None: vDesc_tag = '0.00'
        if icms_st_tag is None: vDesc_tag = '0.00'
        if outros_tag is None: vDesc_tag = '0.00'
        print(f"PRODUTOS {produtos}")
        novo_produtos = {}

        for indice, (chave, valor) in enumerate(produtos.items()):
            cfop = modulos.nbs.de_para_cfop(chave)
            if not cfop: raise ErroNBS(f"Não encontrado CFOP {chave} no de/para")
            cfop_codigo = cfop
            chave_existe = novo_produtos.get(cfop_codigo)
            if chave_existe:
                print("entrei no if")
                novo_produtos[cfop_codigo] += valor
            else:
                print("entrei no else")
                novo_produtos[cfop_codigo] = valor

        #cfop_digito = '1'
        print(novo_produtos)
        #Verificando o cfop de cada produto
        for indice, (chave, valor) in enumerate(novo_produtos.items()):
            print(f"Chave: {chave}")
            print(f"Valor: {valor}")
            cfop_codigo, cfop_digito = chave.split('-')
            valor_decimal = Decimal(valor)
            valor_arredondado = valor_decimal.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            print(f"VALOR ARREDONDADO: {valor_arredondado}")
            #cfop = de_para_cfop(chave)
            # if not cfop: raise ErroNBS(f"Não encontrado CFOP: {chave} no de/para")
            # cfop_codigo = cfop[0]
            # cfop_digito = cfop[1]
            if indice == 0:
                cfop_total = float(valor_arredondado) + float(vIpi_tag) + float(vFrete_tag) + float(icms_st_tag) + float(outros_tag) - float(vDesc_tag)
                #cfop_total_str:str = str(cfop_total)
                cfop_decimal = Decimal(cfop_total)
                cfop_total_arredondado = cfop_decimal.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                bot.logger.informar(cfop_total_arredondado)
                cfop_total_str:str = str(cfop_total_arredondado)
                bot.logger.informar(cfop_total_str)
                cfop_quantidade_total = cfop_total_str.replace(".", ",")
            else:
                cfop_quantidade_total = str(valor_arredondado).replace(".", ",")
            elemento = janela.elementos(class_name='TGroupBox', top_level_only=False)[2]
            elementos: list[HwndWrapper] = elemento.children(class_name='TOvcDbPictureField')[3]
            executar = Elemento(elementos)
            bot.mouse.clicar_mouse(coordenada=executar.coordenada)
            bot.teclado.apertar_tecla('backspace')
            bot.teclado.digitar_teclado(cfop_codigo)
            bot.teclado.apertar_tecla('enter')
            bot.logger.informar(f"Digitei o cfop {cfop_codigo} e dei tab")

            #Procurar imagem de CFOP não existe
            imagem_erro_cfop_nao_existe = bot.imagem.procurar_imagem("imagens/erro_cfop_nao_existe.PNG", confianca=0.8, segundos=5)
            if imagem_erro_cfop_nao_existe:  
                bot.logger.informar("CFOP não existe' ")
                bot.teclado.apertar_tecla('tab')
                raise ErroNBS(f"CFOP não encontrado {cfop_codigo}")  

            sleep(1)
            bot.teclado.digitar_teclado(cfop_digito)
            bot.teclado.apertar_tecla('tab')
            #bot.teclado.apertar_tecla('enter')
            #Procurar imagem de CFOP não existe
            imagem_erro_cfop_nao_existe = bot.imagem.procurar_imagem("imagens/erro_cfop_nao_existe.PNG", confianca=0.8, segundos=5)
            if imagem_erro_cfop_nao_existe:  
                bot.logger.informar("CFOP não existe' ")
                bot.teclado.apertar_tecla('enter')
                raise ErroNBS(f"CFOP não encontrado {cfop_codigo}")  

            #Procurar imagem de Grupo não existe
            imagem_erro_grupo_nao_encontrado = bot.imagem.procurar_imagem("imagens/grupo_nao_encontrado.PNG", confianca=0.8, segundos=5)
            if imagem_erro_grupo_nao_encontrado:  
                bot.logger.informar("Grupo não encontrado para o CFOP' ")
                bot.teclado.apertar_tecla('enter')
                raise ErroNBS(f"Grupo '{cfop_digito} não encontrado para o  '{cfop_codigo}'") 
            
            sleep(1)
            bot.mouse.clicar_mouse(coordenada=executar.coordenada)
            bot.teclado.apertar_tecla('tab')
            bot.teclado.apertar_tecla('tab')
            bot.teclado.apertar_tecla('delete')
            bot.teclado.digitar_teclado(cfop_quantidade_total)
            bot.teclado.apertar_tecla('enter')


            #Clicar em Incluir CFOP
            imagem_incluir_cfop = bot.imagem.procurar_imagem("imagens/incluir_cfop.PNG", confianca=0.8, segundos=5)
            if imagem_incluir_cfop: 
                bot.mouse.clicar_mouse(coordenada=imagem_incluir_cfop) 
                bot.logger.informar("Clicando em 'Incluir CFOP' ")  
            else:
                janela.fechar()
                raise Exception("Não foi possivel clicar em 'Incluir CFOP' ") 

            sleep(1)
            #Procurar Mensagem de Erro
            imagem_erro_cfop = bot.imagem.procurar_imagem("imagens/erro_cfop.PNG", confianca=0.8, segundos=5)
            if imagem_erro_cfop: raise ErroNBS(f"Erro ao inserir codigo Natureza {cfop_codigo} em relação ao CFOP {chave}")      
            else: bot.logger.informar("Não encontrei mensagem de erro, seguindo...") 
        
        #Clicar em Contabilização
        imagem_contabilizacao = bot.imagem.procurar_imagem("imagens/contabilizacao.PNG", confianca=0.8, segundos=5)
        if imagem_contabilizacao: bot.mouse.clicar_mouse(coordenada=imagem_contabilizacao)     
        else: 
            janela.fechar()
            bot.logger.informar("Não foi possivel clicar em Contabilização")
            raise Exception("Não foi possivel clicar em Contabilização")

        #Clicar em Incluir Contabilização
        imagem_incluir_contabilizacao = bot.imagem.procurar_imagem("imagens/incluir_contabilizacao.PNG", confianca=0.8, segundos=5)
        if imagem_incluir_contabilizacao: bot.mouse.clicar_mouse(coordenada=imagem_incluir_contabilizacao)     
        else: 
            janela.fechar()
            bot.logger.informar("Não foi possivel clicar em Incluir Contabilização")
            raise Exception("Não foi possivel clicar em Incluir Contabilização")

        #Aguardar janela de Incluir Conta de Contabilização
        bot.util.aguardar_condicao(lambda: "Incluir Conta de Contabilização" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Incluir Conta de Contabilização")

        #Clicar em Conta Contábil
        elemento = janela.elementos(class_name='TNBSContaContab', top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)


        #Preencher Classificação Contábil
        bot.teclado.digitar_teclado(classificacaoContabil)
        bot.teclado.apertar_tecla('enter')

        sleep(1)
        # Procurar Imagem de erro
        imagem_erro_um = bot.imagem.procurar_imagem("imagens/erro_inserir_classificao_contabil_1.PNG", confianca=0.8, segundos=3)
        if imagem_erro_um: 
            bot.teclado.apertar_tecla('enter')
            raise ErroNBS(f"Erro encontrado ao inserir a classificação contábil: {classificacaoContabil} [FireDAC][Phys][Ora] ORA-01722: invalid number.")     
        else: bot.logger.informar("Não encontrei a imagem, seguindo...")

        imagem_erro_dois = bot.imagem.procurar_imagem("imagens/erro_inserir_classificao_contabil_2.PNG", confianca=0.8, segundos=3)
        if imagem_erro_dois: 
            bot.teclado.apertar_tecla('enter')
            janela.fechar()
            raise ErroNBS(f"Erro encontrado ao inserir a classificação contábil: {classificacaoContabil} Conta Contabil Não Cadastrada ou a Conta é Sintética...")     
        else: bot.logger.informar("Não encontrei a imagem, seguindo")

        #Preencher Centro de Custo
        bot.teclado.digitar_teclado(centroCusto)
        bot.teclado.apertar_tecla('enter')
        #bot.teclado.apertar_tecla('tab')

        # Procurar Imagem de erro
        imagem_erro_centro_custo_um = bot.imagem.procurar_imagem("imagens/erro_centro_custo_1.PNG", confianca=0.8, segundos=3)
        if imagem_erro_centro_custo_um: 
            bot.teclado.apertar_tecla('enter')
            janela.fechar()
            raise ErroNBS(f"Erro encontrado ao inserir o Centro de Custo: {centroCusto} 'Centro de Custo não está ativo.'")     
        else: bot.logger.informar("Não encontrei a imagem, seguindo")

        
        #torre = 'UAB'
        #historicoPadrao = '0' #UAB = 0, Maranhão = 17, Autostar = 25
        historicoPadrao = modulos.nbs.de_para_historico_padrao(torre)

        #Preencher Histórico Padrão
        bot.teclado.apertar_tecla('enter')
        bot.teclado.apertar_tecla('backspace', quantidade=5)
        bot.teclado.apertar_tecla('delete', quantidade=5)
        bot.teclado.digitar_teclado(historicoPadrao)
        bot.logger.informar(f"Preenchido historico Padrão {historicoPadrao}")
        bot.teclado.apertar_tecla('enter', quantidade=4)

        # Procurar Imagem de erro
        imagem_erro_centro_custo_dois = bot.imagem.procurar_imagem("imagens/erro_centro_custo_2.PNG", confianca=0.8, segundos=3)
        if imagem_erro_centro_custo_dois: 
            bot.teclado.apertar_tecla('enter')
            janela.fechar()
            raise ErroNBS(f"Erro encontrado ao inserir o Centro de Custo: {centroCusto} 'Informe o Centro de Custo'")     
        else: bot.logger.informar("Não encontrei a imagem, seguindo")


        #Aguardar janela Entrada Diversas
        bot.util.aguardar_condicao(lambda: "Entrada Diversas / Operação: 52-Entrada Diversas" in bot.windows.Janela.titulos_janelas(), 10)
        janela = bot.windows.Janela("Entrada Diversas / Operação: 52-Entrada Diversas")


        #Clicar em Faturamento
        imagem_faturamento = bot.imagem.procurar_imagem("imagens/faturamento.PNG", confianca=0.8, segundos=5)
        if imagem_faturamento: 
            bot.mouse.clicar_mouse(coordenada=imagem_faturamento) 
            bot.logger.informar("Clicando em Faturamento")    
        else: 
            janela.fechar()
            bot.logger.informar("Não foi possivel clicar em Faturamento")
            raise Exception("Não foi possivel clicar em Faturamento")

        #'TOvcNumericField'
        #Clicar em Total de Parcelas
        elemento = janela.elementos(class_name='TOvcNumericField', top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        bot.teclado.apertar_tecla('backspace')

        #quantidadeParcelas = 1
        bot.logger.informar(f"Digitando quantidade de parcelas: {quantidadeParcelas}")
        bot.teclado.digitar_teclado(str(quantidadeParcelas))
        
        if int(quantidadeParcelas) > 1:
            #Clicar em intervalo
            intervalo = '30'
            elemento = janela.elementos(class_name='TOvcNumericField', top_level_only=False)[1]
            executar = Elemento(elemento)
            bot.mouse.clicar_mouse(coordenada=executar.coordenada)
            bot.teclado.apertar_tecla('backspace')
            bot.logger.informar(f"Digitando intervalo: {intervalo}")
            bot.teclado.digitar_teclado(str(intervalo))

        #Clicar Gerar
        imagem_gerar = bot.imagem.procurar_imagem("imagens/gerar.PNG", confianca=0.8, segundos=5)
        if imagem_gerar: 
            bot.mouse.clicar_mouse(coordenada=imagem_gerar) 
            bot.logger.informar("Clicado em Gerar")    
        else: 
            janela.fechar()
            bot.logger.informar("Não foi possivel clicar em Gerar")
            raise Exception("Não foi possivel clicar em Gerar")

        #Pegando tipo de Pagamento
        tipo_pagamento_natureza_depesa = de_para_tipo_pagamento_natureza_despesa(torre)
        if not tipo_pagamento_natureza_depesa: raise Exception(f"Não foi encontrado tipo de pagamento e natureza de despesa para a torre: {torre}, informada")

        #Clicar em Tipo de Pagamento
        elemento = janela.elementos(class_name='TwwDBLookupCombo', top_level_only=False)[1]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        tipo_pagamento:str = tipo_pagamento_natureza_depesa[0]
        bot.teclado.digitar_teclado(tipo_pagamento)
        bot.logger.informar("Selecionado tipo de Pagamento 'Boleto' ")

        #Clicar em Natureza Despesa
        elemento = janela.elementos(class_name='TwwDBLookupCombo', top_level_only=False)[0]
        executar = Elemento(elemento)
        bot.mouse.clicar_mouse(coordenada=executar.coordenada)
        natureza_depesa:str = tipo_pagamento_natureza_depesa[1]
        bot.teclado.digitar_teclado(natureza_depesa)
        bot.teclado.apertar_tecla('enter')

        # data_vencimento = {
        #     "Data Vencimento 1": "01/07/2024",
        #     "Data Vencimento 2": "01/08/2024",
        #     "Data Vencimento 3": "01/09/2024",
        #     "Data Vencimento 4": "01/10/2024"
        #     }

        #Tratando Parcelamento
        dict_tem_parcelamento = {}
        if temParcelamento.lower() == 'sim':
            for propriedade in propriedades_processo:
                propriedade_value(propriedade, 'Valor Parcela', dict_tem_parcelamento)

                for i in range(1, int(quantidadeParcelas) + 1):
                    propriedade_data(propriedade, f'Data de Vencimento Boleto {i}', dict_tem_parcelamento)
        
        data_formatada = datetime.strptime(dataVencimento, r'%d/%m/%Y')
        data_atual = datetime.now()

        if data_formatada < data_atual:
            nova_data:datetime = adicionar_dia_util(data_atual)
            bot.logger.informar(f"Alterando a  data de Vencimento de {dataVencimento} para {nova_data.strftime(r'%d/%m/%Y')}")
            dataVencimento = nova_data.strftime(r'%d/%m/%Y')
        
        for i in range(1, int(quantidadeParcelas) + 1):
            #'TOvcDbPictureField'
            #Clicar em Data de Vencimento
            elemento = janela.elementos(class_name='TOvcDbPictureField', top_level_only=False)[3]
            executar = Elemento(elemento)
            bot.mouse.clicar_mouse(coordenada=executar.coordenada)
            bot.teclado.apertar_tecla('right', quantidade=10)
            bot.teclado.apertar_tecla('backspace', quantidade=10)
            #bot.teclado.apertar_tecla('delete', quantidade=10)
            if int(quantidadeParcelas) > 1:
                if i == 1: 
                    bot.teclado.digitar_teclado(dataVencimento)
                    bot.logger.informar(f"Digitando vencimento: {dataVencimento}")
                else:
                    bot.teclado.digitar_teclado(f'{dict_tem_parcelamento.get(f'Data de Vencimento Boleto {i}')}')
                    bot.logger.informar(f"Digitando vencimento: {dict_tem_parcelamento.get(f'Data de Vencimento Boleto {i}')} ")
            else:
                bot.teclado.digitar_teclado(f'{dataVencimento}')
                bot.logger.informar(f"Digitado vencimento {dataVencimento}")

            #Clicar em Atualizar Fatura
            imagem_atualizar_fatura = bot.imagem.procurar_imagem("imagens/atualizar_fatura.PNG", confianca=0.8, segundos=5)
            if imagem_atualizar_fatura: 
                bot.mouse.clicar_mouse(coordenada=imagem_atualizar_fatura)
                bot.logger.informar("Clicado em Atualizar Fatura")     
            else: 
                janela.fechar()
                bot.logger.informar("Não foi possivel clicar em Atualizar Fatura")
                raise Exception("Não foi possivel clicar em Atualizar Fatura")


            #Clicar na Seta para mudar de parcela
            imagem_seta = bot.imagem.procurar_imagem("imagens/seta_baixo.PNG", confianca=0.8, segundos=5)
            if imagem_seta: 
                bot.mouse.clicar_mouse(coordenada=imagem_seta)   
                bot.logger.informar("Clicado na seta para mudar parcela")     
            else: 
                janela.fechar()
                bot.logger.informar("Não foi possivel clicar na seta para mudar parcela")
                raise Exception("Não foi possivel clicar na seta para mudar parcela")
            
            # #Clicar em Confirmar
            # imagem_confirmar = bot.imagem.procurar_imagem("imagens/confirmar.PNG", confianca=0.8, segundos=5)
            # if imagem_confirmar: 
            #     bot.mouse.clicar_mouse(coordenada=imagem_confirmar) 
            #     bot.logger.informar("Clicado em confirmar")     
            # else: 
            #     janela.fechar()
            #     bot.logger.informar("Não foi possivel clicar em Confirmar")
            #     raise Exception("Não foi possivel clicar em Confirmar")
            
        sleep(2)
        for i in range(4):
            #Clicar em Confirmar
            imagem_confirmar = bot.imagem.procurar_imagem("imagens/confirmar.PNG", confianca=0.8, segundos=5)
            if imagem_confirmar: 
                bot.mouse.clicar_mouse(coordenada=imagem_confirmar) 
                bot.logger.informar("Clicado em confirmar")    
            else: 
                janela.fechar()
                bot.logger.informar("Não foi possivel clicar em Confirmar")
                raise Exception("Não foi possivel clicar em Confirmar")
            
            #Procurar imagem de vencimento inferior
            imagem_vencimento_inferior = bot.imagem.procurar_imagem(r"imagens/data_vencimento_erro.PNG", confianca=0.8, segundos=3)
            if imagem_vencimento_inferior: 
                bot.teclado.apertar_tecla("enter")
                raise ErroNBS("Alguma data de vencimento inferior a data atual")
            else: bot.logger.informar("Não encontrada a imagem de vencimento inferior, seguindo...")

            #Procurar imagem de interestadual
            imagem_interestadual = bot.imagem.procurar_imagem(r"imagens/entrada_interestadual.PNG", confianca=0.8, segundos=3)
            if imagem_interestadual: 
                bot.logger.informar("Erro de Interestadual, tratando")
                bot.teclado.digitar_teclado("s")
            else: bot.logger.informar("Não encontrada a imagem de interestadual, seguindo...")
            #Procurar imagem de erro de soma de contabilização
            imagem_soma_contabilizacao = bot.imagem.procurar_imagem("imagens/soma_contabilizacao.PNG", confianca=0.8, segundos=3)
            if imagem_soma_contabilizacao: 
                bot.logger.informar("Erro de Soma de Contabilização, tratando")
                bot.teclado.apertar_tecla('enter')
                imagem_alterar_contabilizacao = bot.imagem.procurar_imagem("imagens/alterar_contabilizacao.PNG", confianca=0.8, segundos=5)
                if not imagem_alterar_contabilizacao: raise Exception("Não foi possivel clicar em 'Alterar Contabilização")
                bot.mouse.clicar_mouse(coordenada=imagem_alterar_contabilizacao)
                #Aguardar janela de Incluir Conta de Contabilização
                bot.util.aguardar_condicao(lambda: "Alterar Conta de Contabilização" in bot.windows.Janela.titulos_janelas(), 10)
                janela = bot.windows.Janela("Alterar Conta de Contabilização")
                bot.teclado.apertar_tecla('enter', quantidade=6)
                #Aguardar janela Entrada Diversas
                bot.util.aguardar_condicao(lambda: "Entrada Diversas / Operação: 52-Entrada Diversas" in bot.windows.Janela.titulos_janelas(), 10)
                janela = bot.windows.Janela("Entrada Diversas / Operação: 52-Entrada Diversas")
                continue
            #     #Clicar em Confirmar
            #     imagem_confirmar = bot.imagem.procurar_imagem("imagens/confirmar.PNG", confianca=0.8, segundos=5)
            #     if imagem_confirmar: 
            #         bot.mouse.clicar_mouse(coordenada=imagem_confirmar) 
            #         bot.logger.informar("Clicado em confirmar")
            #         continue     
            #     else: 
            #         janela.fechar()
            #         bot.logger.informar("Não foi possivel clicar em Confirmar")
            #         raise Exception("Não foi possivel clicar em Confirmar")
            # #Procurar imagem de natureza base pis/cofins
            imagem_erro_pis_cofins = bot.imagem.procurar_imagem("imagens/erro_pis_cofins.PNG", confianca=0.8, segundos=3)
            if imagem_erro_pis_cofins: 
                bot.logger.informar("Erro de Pis e Cofins, tratando")
                bot.teclado.apertar_tecla('enter')
                janela = bot.windows.Janela("Entrada Diversas / Operação: 52-Entrada Diversas")
                imagem_capa = bot.imagem.procurar_imagem("imagens/capa.PNG", confianca=0.8, segundos=3)
                if not imagem_capa: raise Exception("Não foi possivel clicar em 'capa'")
                bot.mouse.clicar_mouse(coordenada=imagem_capa)
                elemento = janela.elementos(class_name='TwwDBLookupCombo', top_level_only=False)[1]
                executar = Elemento(elemento)
                bot.mouse.clicar_mouse(coordenada=executar.coordenada)
                texto = 'Aquisicao de bens utilizados como insumo'
                bot.teclado.digitar_teclado(texto)
                bot.teclado.apertar_tecla('tab')
                continue
                    
            else: 
                bot.logger.informar("Não encontrada a imagem de erro pis/cofins, seguindo...")
                        
            #Procurar imagem de erro de Os
            imagem_informar_ordens = bot.imagem.procurar_imagem("imagens/informar_ordens.PNG", confianca=0.8, segundos=3)
            if imagem_informar_ordens: 
                bot.logger.informar("Erro de OS, tratando")
                bot.teclado.apertar_tecla('enter')
                imagem_cadeado = bot.imagem.procurar_imagem("imagens/cadeado.PNG", confianca=0.8, segundos=3)
                if not imagem_cadeado: raise Exception("Não foi localizado o cadeado")
                bot.mouse.clicar_mouse(coordenada=imagem_cadeado)
                imagem_confirmar_os = bot.imagem.procurar_imagem("imagens/confirmacao_os.PNG", confianca=0.8, segundos=3)
                if not imagem_confirmar_os: raise Exception("Não foi localizado a caixa de dialogo de confirmar os, depois de clicar no cadeado")
                bot.teclado.digitar_teclado('S')
                continue
            else:
                bot.logger.informar("Não encontrada a imagem de erro de OS, seguindo...")
                break


                #if i == 3: raise Exception("Erro na entrada da nota após verificar erros de pis/cofins, ")
                        
                # else: 
                #     if i == 2: raise ErroNBS("Erro ao alterar contabilização")
                #     bot.logger.informar("Não encontrada a imagem de erro no total da Nota, seguindo...")
                #     break
            # #Procurar imagem de natureza base pis/cofins
            # imagem_erro_pis_cofins = bot.imagem.procurar_imagem("imagens/erro_pis_cofins.PNG", confianca=0.8, segundos=3)
            # if imagem_erro_pis_cofins: 
            #     bot.teclado.apertar_tecla('enter')
            #     janela.fechar()
            #     raise ErroNBS("Soma de Produtos + IPI + Frete (com base no XML), não é igual ao Total da Nota no NBS")     
            # else: bot.logger.informar("Não encontrada a imagem de erro no total da Nota, seguindo...")
            
        #Procurar imagem de erro de Total de Nota
        imagem_erro_total_nota = bot.imagem.procurar_imagem("imagens/erro-total-nota.PNG", confianca=0.8, segundos=3)
        if imagem_erro_total_nota: 
            bot.teclado.apertar_tecla('enter')
            janela.fechar()
            raise ErroNBS("Soma de Produtos + IPI + Frete (com base no XML), não é igual ao Total da Nota no NBS")     
        else: bot.logger.informar("Não encontrada a imagem de erro no total da Nota, seguindo...")

        #Procurar imagem de já lançada
        # imagem_nota_lancada = bot.imagem.procurar_imagem("imagens/nota_lancada.PNG", confianca=0.8, segundos=3)
        # if imagem_nota_lancada: 
        #     bot.teclado.apertar_tecla('enter')
        #     janela.fechar()
        #     raise ErroNBS("Nota ja Lançada")     
        # else: bot.logger.informar("Não encontrada a imagem de já lançada, seguindo...")

        janela_informacao = bot.util.aguardar_condicao(lambda: "Informação" in bot.windows.Janela.titulos_janelas(), 10)
        if janela_informacao:
            bot.teclado.atalho_teclado(["ctrl", "c"])
            texto: str = bot.teclado.texto_copiado(True)
            if "número da nota fiscal já lançada" in texto.lower():
                raise ErroNBS("Nota ja Lançada")
            
        sleep(1)

        #Clicar em Ok
        imagem_sucesso = bot.imagem.procurar_imagem("imagens/sucesso_entrada.PNG", confianca=0.8, segundos=40)
        if imagem_sucesso: 
            bot.teclado.apertar_tecla('enter')
            bot.logger.informar("Clicando em Ok")    
        else: 
            raise ErroNBS("Erro após confirmar a nota, favor verificar")
        
        #Aguardar Janela de Controle
        bot.util.aguardar_condicao(lambda: "Ficha de Controle de Pagamento" in bot.windows.Janela.titulos_janelas(), 60)
        janela = bot.windows.Janela("Ficha de Controle de Pagamento")
        sleep(1)
        janela.fechar()

        bot.logger.informar("Fim do Processamento da tarefa")
        return RetornoStatus(True)


    except ErroNBS as erro:
        bot.logger.alertar(f"Erro no processo de entrada de nota no NBS: { erro }")
        return RetornoStatus(False, f"Erro no processo de entrada de nota: { erro }")
    
    except Exception as erro:
        bot.logger.alertar(f"Erro não mapeado no processo de entrada de nota: { erro }")
        return RetornoStatus(False, f"Erro no processo de entrada de nota: { erro }")

def de_para_cfop(cfop:str) -> str:
    """Função para de/para no preenchimento do CFOP"""
    dict_cfop:dict = {
        "5101": ("1556-1"),
        "5655": ("1652-1"),
        "5102": ("1556-1"),
        "6101": ("2556-1"),
        "6102": ("2556-1"),
        "5401": ("1407-1"),
        "5403": ("1407-1"),
        "6401": ("2407-1"),
        "6403": ("2407-1"),
        "5405": ("1407-1"),
        "5656": ("1653-1"),
        "5929": ("1556-1"),
        "6929": ("2556-1"),
        "6656": ("2653-1")     
    }
    return dict_cfop.get(cfop)

def produtos_xml() -> dict | Exception :
    """Função para fazer a leitura dos produtos do XML"""

    nfe_xml = 'tmp_nfe.xml'
    try:
        with open(nfe_xml, 'r', encoding='utf-8') as nfe:
            xml = ET.parse(nfe)
            raiz = xml.getroot()

            namespace = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            #Localizar os produtos no XML
            det_tags = raiz.findall('.//nfe:det', namespace)
            #Iterar pelas tags encontradas e extrair os produtos
            produtos = {}
            for det in det_tags:
                prod = det.find('nfe:prod', namespace)
                if prod is not None:
                    cProd = prod.find('nfe:CFOP', namespace).text if prod.find('nfe:CFOP', namespace) is not None else None
                    vProd = float(prod.find('nfe:vProd', namespace).text) if prod.find('nfe:vProd', namespace) is not None else None
                    chave_existe = produtos.get(cProd)
                    if chave_existe:
                        print("entrei no if")
                        produtos[cProd] += vProd
                    else:
                        print("entrei no else")
                        produtos[cProd] = vProd
    except Exception as e:
        bot.logger.informar(f"Erro na leitura do XML {e}")
        raise Exception(f"Erro: {e}")


    return produtos

def valores_xml():
    "Função para pegar os valores do XML"
    nfe_xml = 'tmp_nfe.xml'
    try:
        with open(nfe_xml, 'r', encoding='utf-8') as nfe:
            xml = ET.parse(nfe)
            raiz = xml.getroot()

            namespace = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            #Localizar no XML os valores totais
            vIpi_tag = raiz.find('.//nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vIPI', namespace)
            vFrete_tag = raiz.find('.//nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vFrete', namespace)
            vDesc_tag = raiz.find('.//nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vDesc', namespace)
            icms_st_tag = raiz.find('.//nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vST', namespace)
            outros_tag = raiz.find('.//nfe:infNFe/nfe:total/nfe:ICMSTot/nfe:vOutro', namespace)
            serie_tag = raiz.find('.//nfe:infNFe/nfe:ide/nfe:serie', namespace)           
    except Exception as e:
        bot.logger.informar(f"Erro na leitura do XML {e}")
        raise Exception(f"Erro: {e}")
    
    return vIpi_tag.text, vFrete_tag.text, vDesc_tag.text, serie_tag.text, icms_st_tag.text, outros_tag.text 

def de_para_historico_padrao(torre:str) -> str:
    "Função de de/para para o historico padrão a partir da Torre"
    historico_padrao:dict = {
        "UAB": "0",
        "Maranhão": "17",
        "Original": "17",
        "Autostar": "25"
    }

    return historico_padrao.get(torre)

def de_para_tipo_pagamento_natureza_despesa(torre:str):
    "Função para pegar o Tipo de Pagamento e Despesa a partir da Torre"
    tipo_pagamento_natureza_desposa:dict = {
        "UAB": ("Boleto", "OUTRAS DESPESAS"),
        "Autostar": ("Dinheiro", "Despesas Administrativas"),
        "Maranhão": ("Carteira", "Adiant. Funcionarios"),
        "Original": ("Carteira", "Adiant. Funcionarios")
    }
    return tipo_pagamento_natureza_desposa.get(torre)

def propriedade_value(propriedade:str, nome:str, lista:dict):
    "Função para pegar a propriedade e atribuir o value dela em um novo dict"
    if propriedade['name'].lower() == nome.lower():
        valor = propriedade.get('value')
        lista[nome] = valor
def propriedade_data(propriedade:str, nome:str, lista:dict):
    "Função para pegar a propriedade e atribuir o value dela em um novo dict formatando a data"
    if propriedade['name'].lower() == nome.lower():
        valor = propriedade.get('value')
        if valor: lista[nome] = datetime.fromisoformat(valor).strftime(r'%d/%m/%Y')
        else: lista[nome] = valor

# def encerrar_sistema():
#     """Encerra o sistema NBS"""
#     processos = bot.configfile.obter_opcao_ou("nbs", "processos")

#     comando = f"""
#         Get-Process -Name {processos} -IncludeUserName |
#             Where-Object {{ $_.UserName -eq "$env:USERDOMAIN\\$env:USERNAME" }} |
#             Stop-Process -Force
#     """
#     subprocess.run(["powershell", "-Command", comando], check=False)

def comparar_data():
    "Função para comparar se a data atual é maior que 28 do mês a 12:00"
    dia = 28
    hora_comparacao = 12
    minuto_comparacao = 0

    #Data e hora atual
    momento_atual = datetime.now()

    #Data e hora de comparação
    data_comparacao = momento_atual.replace(day=dia, hour=hora_comparacao, minute=minuto_comparacao, second=0, microsecond=0)

    #Verificar se o momento atual é maior ou igual a data de comparação
    if momento_atual >= data_comparacao:
        proximo_mes = momento_atual.replace(day=1) + timedelta(days=32)
        data = proximo_mes.replace(day=1)
        data_formatada = data.strftime(r"%d/%m/%Y")
    else:
        data_formatada = None

    return data_formatada

def adicionar_dia_util(data:datetime) -> datetime:
    """Função para adicionar um dia útil para o vencimento"""
    ano_atual = data.year
    feriados_brasil = holidays.Brazil(years=ano_atual)

    while True:
        data+= timedelta(days=1)

        #Verifica se o ano mudou
        if data.year != ano_atual:
            ano_atual = data.year
            #Atualiza a lista de feriados para o novo ano
            feriados_brasil = holidays.Brazil(years=ano_atual)

        #Verifica se é dia util (segunda a sexta) e não é feriado
        if data.weekday() < 5 and data not in feriados_brasil:
            break
    return data

__all__ = [
    # "iniciar_sistema",
    "ErroNBS",
    "Sistema",
    "processar_entrada",
    "produtos_xml",
    "consultar_de_para_empresa",
    "de_para_cfop",
    "de_para_historico_padrao",
    # "encerrar_sistema",
    "comparar_data",
    "valores_xml"
]