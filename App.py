import streamlit as st
from PyPDF2 import PdfReader
from io import StringIO
import csv
import unicodedata

# Exportar para excel 
import pandas as pd
from io import BytesIO

data_linhas = []
######################

from funcoes import DadosRetornoCSV  # sua função existente

st.set_page_config(page_title="Extrator COELBA", layout="wide")

st.title("⚡ Extrator Inteligente de Faturas PDF. Versão: 1.2")

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
    type="pdf",
    accept_multiple_files=True
)

def normalizar_texto(txt):
    txt = unicodedata.normalize("NFKD", txt)
    txt = txt.replace("º", "o").replace("°", "o").replace("ᵒ", "o")
    #print('Início: ', txt.find('Conta  Contrato Coletiva no'))
    return txt

if uploaded_files:

    output = StringIO(newline='')

    writer = csv.writer(output, quoting=csv.QUOTE_ALL, delimiter=';')

    # Cabeçalho (igual ao seu)
    lista_cabecalho = [
        'Dados do cliente', 
        'Endereço da Unidade Consumidora', 
        'Número da Nota Fiscal', 
        'N da Instalação', 
        'Classificação', 
        'Descrição da Nota Fiscal', 
        'Tarifas Aplicadas', 
        'ICMS Base de Cálculo', 
        'ICMS Base 2', 
        'ICMS Base 3', 
        'ICMS Base 4', 
        'ICMS Base 5', 
        'ICMS Base 6', 
        'ICMS Base 7', 
        'ICMS Base 8', 
        'ICMS Base 9', 
        'Número do medidor', 
        'Conta Contrato', 
        'Mês Ano', 
        'Total a pagar'
    ]

    writer.writerow(lista_cabecalho)

    total_boletos = 0
    lista_mes_ano_aux_concat = ""

    progress = st.progress(0)
    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):

        st.write(f"📄 Processando: {file.name}")

        reader = PdfReader(file)

        contPag = 0
        controlarPag = 0

        # Trecho responsável por identificar o ano de refererência do boleto.
        # Verificar o ano do boleto
        # Mês Ano
        ano2022 = False
        ano2023 = False
        ano2024 = False
        ano2025 = False
        ano2026 = False
        
        pageAux = reader.pages
        for pageAux in reader.pages:
          texto_aux = pageAux.extract_text()
          if not texto_aux:
              continue

          if texto_aux.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and \
             texto_aux.find('AVISO IMPORTANTE!') == -1 and \
             texto_aux.find('2 /   3') == -1 and \
             texto_aux.find('3 /   3') == -1:

             lista_mes_ano_aux = ""
             verificarMesAux = '99'

             TEXTO_COMPLETO_AUX = texto_aux

            #  print('TEXTO_COMPLETO_AUX: ', TEXTO_COMPLETO_AUX)
             
            #  print(TEXTO_COMPLETO_AUX)
            #  print("=================================")
            #  print("")             

             if texto_aux.find('AUTENTICAÇÃO MECÂNICA') > 0:
               lista_mes_ano_aux = DadosRetornoCSV(len('MÊS/ANO'), texto_aux.find('MÊS/ANO'), texto_aux.find('TOTAL A PAGAR(R$)')+1, TEXTO_COMPLETO_AUX) 
               lista_mes_ano_aux_concat = lista_mes_ano_aux
             elif texto_aux.find('1 /   3') > 0:
               lista_mes_ano_aux = DadosRetornoCSV(len('DATA DA EMISSÃO DA NOTA FISCAL'), texto_aux.find('DATA DA EMISSÃO DA NOTA FISCAL')+4, texto_aux.find('DATA DA APRESENTAÇÃO')+1, TEXTO_COMPLETO_AUX) 
               lista_mes_ano_aux_concat = DadosRetornoCSV(len('DATA DA EMISSÃO DA NOTA FISCAL'), texto_aux.find('DATA DA EMISSÃO DA NOTA FISCAL'), texto_aux.find('DATA DA APRESENTAÇÃO')+1, TEXTO_COMPLETO_AUX) 
             else:
               lista_mes_ano_aux = DadosRetornoCSV(len('REF:MÊS/ANO'), texto_aux.find('REF:MÊS/ANO')+3, texto_aux.find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 
               lista_mes_ano_aux_concat = DadosRetornoCSV(len('REF:MÊS/ANO'), texto_aux.find('REF:MÊS/ANO'), texto_aux.find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 
               verificarMes = DadosRetornoCSV(len('REF:MÊS/ANO'), texto_aux.find('REF:MÊS/ANO'), texto_aux.find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX)
               verificarMesAux = verificarMes[0:2] 
       
             if verificarMesAux == '99':
               verificarMesAux = lista_mes_ano_aux[0:2]

            #  print('verificarMes: ', verificarMes)

             #if (lista_mes_ano_aux[2:7] == '/2022' or lista_mes_ano_aux == '/2022' or lista_mes_ano_aux[2:7] == '/2021' or lista_mes_ano_aux == '/2021') and not (int(verificarMes[0:2]) == 9 and (lista_mes_ano_aux[2:7] == '/2022' or lista_mes_ano_aux == '/2022')): 
             if ((lista_mes_ano_aux[2:7] == '/2022' or lista_mes_ano_aux == '/2022') and int(verificarMesAux) < 9 ) or lista_mes_ano_aux[2:7] == '/2021' or lista_mes_ano_aux == '/2021':              
                ano2022 = True 
                break 
             else:
             #lista_mes_ano_aux_2024 = DadosRetornoCSV(len('MÊS/ANO'), pageAux.extract_text().find('MÊS/ANO'), pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 
             
             #if lista_mes_ano_aux_2024[2:7] == '/2026' or lista_mes_ano_aux_2024[2:7] == '/2025' or lista_mes_ano_aux_2024[2:7] == '/2024' or lista_mes_ano_aux_2024[2:7] == '/2023' or (9 >= int(verificarMes[0:2]) < 13 and (lista_mes_ano_aux[2:7] == '/2022' or lista_mes_ano_aux == '/2022')):
                ano2024 = True 
                break

        # print('Ano 2022: ',ano2022)
        # print('Ano 2024: ',ano2024)
        # print('lista_mes_ano_aux: ', lista_mes_ano_aux)
        # print('lista_mes_ano_aux[2:7]: ', lista_mes_ano_aux[2:7])
        # print('lista_mes_ano_aux: ', lista_mes_ano_aux)
        # print('verificarMes[0:2]: ', lista_mes_ano_aux[0:2])                

        if ano2024:
            for page in reader.pages:
             
                texto = page.extract_text()
                if texto is None:
                    continue

                if (texto.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    texto.find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    texto.find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    texto.find('DIC, FIC, DMIC e DICRI') == -1) :
                           
                    TEXTO_COMPLETO = texto 

                    # print(TEXTO_COMPLETO)

                    try:
                        lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), texto.find('NOME DO CLIENTE:'), texto.find('ENDEREÇO:'), TEXTO_COMPLETO)

                        lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), texto.find('ENDEREÇO:'), texto.find('CÓDIGO DA')+1, TEXTO_COMPLETO)

                        lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), texto.find('NOTA FISCAL N°'), texto.find('- SÉRIE'), TEXTO_COMPLETO)

                        lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), texto.find('INSTALAÇÃO'), texto.find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)

                        lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), texto.find('CLASSIFICAÇÃO:'), texto.find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)

                        if texto.find('neoenergiacoelba.com.br') > 0:
                            lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
                        else:
                            lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('www.neoenergia.com'), TEXTO_COMPLETO)

                        lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split())
                        lista_desc_nota_fiscal_tratado_list = lista_desc_nota_fiscal_tratado.split(" ")

                        lista_desc_nota_fiscal_gerar = " ".join(lista_desc_nota_fiscal_tratado_list)

                        # Tarifas
                        lista_desc_tarifa_gerar = " ".join([
                            lista_desc_nota_fiscal_tratado_list[0],
                            lista_desc_nota_fiscal_tratado_list[9],
                            lista_desc_nota_fiscal_tratado_list[10],
                            lista_desc_nota_fiscal_tratado_list[19],
                        ])

                        # ICMS
                        lista_inform_tributos_ICMS = DadosRetornoCSV(len('ICMS'), texto.find('ICMS'), texto.find('CONSUMO / kWh'), TEXTO_COMPLETO)
                        lista_inform_tributos_list_ICMS = " ".join(lista_inform_tributos_ICMS.split()).split(" ")

                        lista_inform_tributos_PIS = DadosRetornoCSV(len('(%) PIS'), texto.find('(%) PIS'), texto.find('COFINS'), TEXTO_COMPLETO)
                        lista_inform_tributos_list_PIS = " ".join(lista_inform_tributos_PIS.split()).split(" ")

                        lista_inform_tributos_COFINS = DadosRetornoCSV(len('COFINS'), texto.find('COFINS'), texto.find('ICMS'), TEXTO_COMPLETO)
                        lista_inform_tributos_list_COFINS = " ".join(lista_inform_tributos_COFINS.split()).split(" ")

                        if texto.find('MEDIDOR kWh') > 0 and texto.find('Energia Ativa') > 0:
                            lista_num_medidor = DadosRetornoCSV(len('MEDIDOR kWh'), texto.find('MEDIDOR kWh'), texto.find('Energia Ativa'), TEXTO_COMPLETO)
                        else:
                            lista_num_medidor = ""                                                

                        if texto.find('CÓDIGO DO CLIENTE') > 0 and texto.find('DATAS DE LEITURAS  LEITURA ANTERIOR') > 0:
                          print('DATAS DE LEITURAS  LEITURA ANTERIOR')
                          lista_conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), texto.find('CÓDIGO DO CLIENTE'), texto.find('DATAS DE LEITURAS'), TEXTO_COMPLETO)
                        
                        # if texto.find(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE') != -1 and texto.find('VENCIMENTO') != -1:                                                   
                        #     lista_conta_contato = DadosRetornoCSV(len(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE'), texto.find(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE'), texto.find('VENCIMENTO')+12, TEXTO_COMPLETO)
                            
                        #     print('VENCIMENTO')  
                        #     print('lista_conta_contato: ', lista_conta_contato)
                        #     print(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE')
                        #     print('len(lista_mes_ano_aux_concat +  CÓDIGO DO CLIENTE)', len(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE'))

                        #     print('texto.find(lista_mes_ano_aux_concat + CÓDIGO DO CLIENTE): ', texto.find(lista_mes_ano_aux_concat + ' CÓDIGO DO CLIENTE') )
                        #     print('VENCIMENTO+12', texto.find('VENCIMENTO')+12)

                        # elif texto.find('CÓDIGO DO CLIENTE') > 0 and texto.find('DATAS DE LEITURAS  LEITURA ANTERIOR') > 0:
                        #     print('DATAS DE LEITURAS  LEITURA ANTERIOR')
                        #     lista_conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), texto.find('DATAS DE LEITURAS  LEITURA ANTERIOR'), texto.find('DATAS DE LEITURAS  LEITURA ANTERIOR'), TEXTO_COMPLETO)
                        elif texto.find('Conta  Contrato Coletiva nº') > 0:
                            if texto.find('Regras para cobrança da contribuição para o custeio do serviço de') > 0:
                                lista_conta_contato = DadosRetornoCSV(len('Conta  Contrato Coletiva nº'), texto.find('Conta  Contrato Coletiva nº'), texto.find('Regras para cobrança da contribuição para o custeio do serviço de'), TEXTO_COMPLETO).replace(".", "")                                
                            elif texto.find('A partir de agosto o IBGE realizará o censo demográfico 2022') > 0:
                                lista_conta_contato = DadosRetornoCSV(len('Conta  Contrato Coletiva nº'), texto.find('Conta  Contrato Coletiva nº'), texto.find('A partir de agosto o IBGE realizará o censo demográfico 2022'), TEXTO_COMPLETO).replace(".", "")
                            else:
                                lista_conta_contato = DadosRetornoCSV(len('Conta  Contrato Coletiva nº'), texto.find('Conta  Contrato Coletiva nº'), texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO) .replace(".", "")                                                               
                        else:
                            if texto.find('A partir de agosto o IBGE realizará o censo demográfico 2022') > 0:                               
                                lista_conta_contato = DadosRetornoCSV(len('Conta  Contrato Coletiva nº'), texto.find('Conta  Contrato Coletiva nº'), texto.find('A partir de agosto o IBGE realizará o censo demográfico 2022'), TEXTO_COMPLETO).replace(".", "")
                            else:
                                lista_conta_contato = DadosRetornoCSV(len('Conta Contrato Coletiva nº'), texto.find('Conta Contrato Coletiva nº'), texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO).replace(".", "")                                                                                                                                                                                        

                        lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto.find('MÊS/ANO'), texto.find('VENCIMENTO')+1, TEXTO_COMPLETO)

                        lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto.find('TOTAL A PAGAR R$'), texto.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                        data_linhas.append([ lista_dados_cliente,
                            lista_end_unid_consum,
                            lista_num_nota_fiscal,
                            lista_num_Instalacao,
                            lista_classificacao,
                            lista_desc_nota_fiscal_gerar,
                            lista_desc_tarifa_gerar,
                            lista_inform_tributos_list_ICMS[0],
                            lista_inform_tributos_list_ICMS[1],
                            lista_inform_tributos_list_ICMS[2],
                            lista_inform_tributos_list_PIS[0],
                            lista_inform_tributos_list_PIS[1],
                            lista_inform_tributos_list_PIS[2],
                            lista_inform_tributos_list_COFINS[0],
                            lista_inform_tributos_list_COFINS[1],
                            lista_inform_tributos_list_COFINS[2],
                            lista_num_medidor,
                            lista_conta_contato,
                            lista_mes_ano,
                            lista_total_pagar
                        ])

                        total_boletos += 1
                        controlarPag += 1

                    except Exception as e:
                        st.warning(f"Erro ao processar página: {e}")

                elif texto.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
                    controlarPag = 0

            contPag += 1
        elif ano2022:
           for page in reader.pages: 
                          
              texto = page.extract_text()              
              
              if texto is None:
                    continue                                    
              
              if (texto.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and
                  texto.find('AVISO IMPORTANTE!') == -1 and
                  texto.find('2 /   3') == -1 and
                  texto.find('3 /   3') == -1):
         
                  TEXTO_COMPLETO = texto

                  #print('Entrei em 2022.') 

                  try:                                         

                    # Dados do cliente
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_dados_cliente = DadosRetornoCSV(len('DADOS DO CLIENTE'), texto.find('DADOS DO CLIENTE'), texto.find('DATA DE VENCIMENTO'), TEXTO_COMPLETO)            
                    elif texto.find('1 /   3') > 0:
                        lista_dados_cliente = DadosRetornoCSV(len('DADOS DO CLIENTE'), texto.find('DADOS DO CLIENTE'), texto.find('ENDEREÇO'), TEXTO_COMPLETO)

                    # print('lista_dados_cliente: ', lista_dados_cliente)

                    # Endereço Unidade Consumidora
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO DA UNIDADE CONSUMIDORA'), texto.find('ENDEREÇO DA UNIDADE CONSUMIDORA'), texto.find('RESERVADO AO FISCO'), TEXTO_COMPLETO)                      
                    elif texto.find('1 /   3') > 0:
                        lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO'), texto.find('ENDEREÇO'), texto.find('DATA DE VENCIMENTO')+1, TEXTO_COMPLETO) 

                    # Número da Nota Fiscal
                    lista_num_nota_fiscal = DadosRetornoCSV(len('NÚMERO DA NOTA FISCAL'), texto.find('NÚMERO DA NOTA FISCAL'), texto.find('CONTA CONTRATO'), TEXTO_COMPLETO)          
                
                    # Nº da Instlação
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_num_Instalacao = DadosRetornoCSV(len('Nº DA INSTALAÇÃO'), texto.find('Nº DA INSTALAÇÃO'), texto.find('CLASSIFICAÇÃO'), TEXTO_COMPLETO)                      
                    elif texto.find('1 /   3') > 0:
                        lista_num_Instalacao = DadosRetornoCSV(len('Nº DA INSTALAÇÃO'), texto.find('Nº DA INSTALAÇÃO'), texto.find('RESERVADO AO FISCO'), TEXTO_COMPLETO)
                    
                    # Classificação
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO'), texto.find('CLASSIFICAÇÃO'), texto.find('ENDEREÇO DA UNIDADE CONSUMIDORA'), TEXTO_COMPLETO)          
                    elif texto.find('1 /   3') > 0:
                        lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO'), texto.find('CLASSIFICAÇÃO'), texto.find('DESCRIÇÃO DA NOTA FISCAL E INFORMAÇÕES IMPORTANTES'), TEXTO_COMPLETO)          

                    # Descrição da Nota Fiscal
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('DESCRIÇÃO DA NOTA FISCAL'), TEXTO_COMPLETO)
                    elif texto.find('1 /   3') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('DESCRIÇÃO DA NOTA FISCAL E INFORMAÇÕES IMPORTANTES'), TEXTO_COMPLETO)
                    elif texto.find('REF:MÊS/ANO') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
                                                                                                                    
                    lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split())

                    # Tarifas Aplicadas
                    if texto.find('DATA PREVISTA DA PRÓXIMA LEITURA:') > 0:
                        lista_tarifas_aplicadas_tratada = DadosRetornoCSV(len('DATA PREVISTA DA PRÓXIMA LEITURA:')+11, texto.find('DATA PREVISTA DA PRÓXIMA LEITURA:'), texto.find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
                    elif texto.find('1 /   3') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto.find('NOTA FISCAL | FATURA | CONTA DE ENERGIA ELÉTRICA'), TEXTO_COMPLETO)
                    else:
                        lista_tarifas_aplicadas = DadosRetornoCSV(len('AJUSTECONSUMO'), texto.find('AJUSTECONSUMO'), texto.find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
                        lista_tarifas_aplicadas_tratada = lista_tarifas_aplicadas.replace('(kWh)','').strip()

                    # Informações de Tributos
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_inform_tributos = DadosRetornoCSV(len('INFORMAÇÕES DE TRIBUTOS'), texto.find('INFORMAÇÕES DE TRIBUTOS'), texto.find('AUTENTICAÇÃO MECÂNICA'), TEXTO_COMPLETO)
                    elif texto.find('1 /   3') > 0:
                        lista_inform_tributos = DadosRetornoCSV(len('CÁLCULO % IMPOSTO PIS/COFINS % IMPOSTO % IMPOSTO'), texto.find('CÁLCULO % IMPOSTO PIS/COFINS % IMPOSTO % IMPOSTO'), texto.find('1 /   3'), TEXTO_COMPLETO)
                    
                    lista_inform_tributos_tratado = " ".join(lista_inform_tributos.split()) # Retrar os espaços entre as palavras      
                    lista_inform_tributos_list = lista_inform_tributos_tratado.split(" ") # Converter em lista  

                    diferencaTotal = -1
                    lista_inform_tributos_list_aux = [' ',' ',' ',' ',' ',' ',' ',' ',' ']

                    for x in range(len(lista_inform_tributos_list)):                               
                        lista_inform_tributos_list_aux[diferencaTotal] = lista_inform_tributos_list[diferencaTotal]
                        diferencaTotal -=1                      

                    # Número do Medidor
                    if texto.find('CAT') > 0 and texto.find('AJUSTECONSUMO') > 0:
                        lista_num_medidor = DadosRetornoCSV(len('AJUSTECONSUMO'), texto.find('AJUSTECONSUMO'), texto.find('CAT'), TEXTO_COMPLETO)
                        lista_num_medidor_tratado = lista_num_medidor.replace('(kWh)','').strip()
                    else:
                        lista_num_medidor_tratado = ""

                    # Conta Contrato
                    lista_conta_contato = DadosRetornoCSV(len('CONTA CONTRATO'), texto.find('CONTA CONTRATO'), texto.find('Nº DO CLIENTE'), TEXTO_COMPLETO)
                
                    # Mês Ano
                    if texto.find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto.find('MÊS/ANO'), texto.find('TOTAL A PAGAR(R$)')+1, TEXTO_COMPLETO) 
                    elif texto.find('1 /   3') > 0:
                        lista_mes_ano = DadosRetornoCSV(len('DATA DA EMISSÃO DA NOTA FISCAL'), texto.find('DATA DA EMISSÃO DA NOTA FISCAL')+4, texto.find('DATA DA APRESENTAÇÃO')+1, TEXTO_COMPLETO) 
                    else:
                        lista_mes_ano = DadosRetornoCSV(len('REF:MÊS/ANO'), pageAux.extract_text().find('REF:MÊS/ANO')+3, pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 

                    # Total a pagar
                    lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR (R$)'), texto.find('TOTAL A PAGAR (R$)'), texto.find('DATA DA EMISSÃO DA NOTA FISCAL'), TEXTO_COMPLETO)                  

                    data_linhas.append([
                        lista_dados_cliente, 
                        lista_end_unid_consum, 
                        lista_num_nota_fiscal, 
                        lista_num_Instalacao, 
                        lista_classificacao, 
                        lista_desc_nota_fiscal_tratado, 
                        lista_tarifas_aplicadas_tratada, 
                        lista_inform_tributos_list_aux[0], 
                        lista_inform_tributos_list_aux[1],
                        lista_inform_tributos_list_aux[2],
                        lista_inform_tributos_list_aux[3],
                        lista_inform_tributos_list_aux[4],
                        lista_inform_tributos_list_aux[5],
                        lista_inform_tributos_list_aux[6],
                        lista_inform_tributos_list_aux[7],
                        lista_inform_tributos_list_aux[8],
                        lista_num_medidor_tratado,
                        lista_conta_contato,
                        lista_mes_ano,
                        lista_total_pagar
                    ])

                    total_boletos += 1
                    controlarPag += 1

                  except Exception as e:
                        st.warning(f"Erro ao processar página: {e}")
        contPag += 1
        progress.progress((i + 1) / total_files)        

    # CSV final
    csv_string = output.getvalue()
    csv_bytes = csv_string.encode("utf-8-sig")

    st.success(f"✅ Total de boletos extraídos: {total_boletos}")

    ## Exportar para excel 

    output_excel = BytesIO()

    df = pd.DataFrame(data_linhas, columns=lista_cabecalho)

    colunas_moeda = [
        "Total a pagar",
        "ICMS Base de Cálculo",
        "ICMS Base 2",
        "ICMS Base 3",
        "ICMS Base 4",
        "ICMS Base 5",
        "ICMS Base 6",
        "ICMS Base 7",
        "ICMS Base 8",
        "ICMS Base 9"
    ]

    def converter_moeda(valor):
        try:
            return float(valor.replace('.', '').replace(',', '.'))
        except:
            return valor

    for col in colunas_moeda:
        if col in df.columns:
            df[col] = df[col].apply(converter_moeda)

    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório')

        workbook  = writer.book
        worksheet = writer.sheets['Relatório']

        # Número de linhas e colunas
        (max_row, max_col) = df.shape

        # Criar tabela com estilo
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': [{'header': col} for col in df.columns],
            'style': 'Table Style Medium 9'  # ← você pode trocar o estilo
        })

        # ❄️ Congelar cabeçalho
        worksheet.freeze_panes(1, 0)

        # 💰 Formato de moeda (R$)
        formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # 👉 Ajustar largura das colunas automaticamente
        for col_num, col_name in enumerate(df.columns):
            max_length = max(
                df[col_name].astype(str).map(len).max(),
                len(col_name)
            ) + 2

            worksheet.set_column(col_num, col_num, max_length)

            # 💰 Aplicar moeda apenas na coluna "Total a pagar"

            if col_name in colunas_moeda:
                worksheet.set_column(col_num, col_num, max_length, formato_moeda)                

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'middle',
            'align': 'center'
        })

    excel_bytes = output_excel.getvalue()

    st.download_button(
    label="📥 Baixar Excel Profissional",
    data=excel_bytes,
    file_name="relatorio_coelba.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
 