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

st.title("⚡ Extrator Inteligente de Faturas PDF")

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
          if pageAux.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and \
             pageAux.extract_text().find('AVISO IMPORTANTE!') == -1 and \
             pageAux.extract_text().find('2 /   3') == -1 and \
             pageAux.extract_text().find('3 /   3') == -1:

             lista_mes_ano_aux = ""

             TEXTO_COMPLETO_AUX = pageAux.extract_text()
             
            #  print(TEXTO_COMPLETO_AUX)
            #  print("=================================")
            #  print("")

             if pageAux.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
               lista_mes_ano_aux = DadosRetornoCSV(len('MÊS/ANO'), pageAux.extract_text().find('MÊS/ANO'), pageAux.extract_text().find('TOTAL A PAGAR(R$)')+1, TEXTO_COMPLETO_AUX) 
             elif pageAux.extract_text().find('1 /   3') > 0:
               lista_mes_ano_aux = DadosRetornoCSV(len('DATA DA EMISSÃO DA NOTA FISCAL'), pageAux.extract_text().find('DATA DA EMISSÃO DA NOTA FISCAL')+4, pageAux.extract_text().find('DATA DA APRESENTAÇÃO')+1, TEXTO_COMPLETO_AUX) 
             else:
               lista_mes_ano_aux = DadosRetornoCSV(len('REF:MÊS/ANO'), pageAux.extract_text().find('REF:MÊS/ANO')+3, pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 

            #  print('Mês/Ano: ', lista_mes_ano_aux)                 
            #  print('REF:MÊS/ANO: ', DadosRetornoCSV(len('REF:MÊS/ANO'), pageAux.extract_text().find('REF:MÊS/ANO')+3, pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) )

            #  print('lista_mes_ano_aux[2:7]:', lista_mes_ano_aux[2:7])

             if lista_mes_ano_aux[2:7] == '/2022' or lista_mes_ano_aux == '/2022' or lista_mes_ano_aux[2:7] == '/2021' or lista_mes_ano_aux == '/2021':
               ano2022 = True
               break
             
             # print('lista_mes_ano_aux[2:7]: ', lista_mes_ano_aux[2:7])

             lista_mes_ano_aux_2024 = DadosRetornoCSV(len('MÊS/ANO'), pageAux.extract_text().find('MÊS/ANO'), pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX)              

             if lista_mes_ano_aux_2024[2:7] == '/2026' or lista_mes_ano_aux_2024[2:7] == '/2025' or lista_mes_ano_aux_2024[2:7] == '/2024' or lista_mes_ano_aux_2024[2:7] == '/2023':
               ano2024 = True
               break
                                         
        if ano2024:
            for page in reader.pages:
             
                texto = page.extract_text()

                if texto is None:
                    continue

                if (page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    page.extract_text().find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    page.extract_text().find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    page.extract_text().find('DIC, FIC, DMIC e DICRI') == -1) :
                           
                    TEXTO_COMPLETO = texto 

                    #print('Seguindo: ', TEXTO_COMPLETO.find('Conta  Contrato Coletiva no'))
                    # print('texto.find(Conta Contrato Coletiva: )', texto.find('Conta Contrato Coletiva no'))

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

                        #textoContaContrato = normalizar_texto(texto)
                        if page.extract_text().find('Conta  Contrato Coletiva nº') > 0:
                            lista_conta_contato = DadosRetornoCSV(len('Conta  Contrato Coletiva nº'), texto.find('Conta  Contrato Coletiva nº'), texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO) .replace(".", "")
                            #lista_conta_contato = normalizar_texto(lista_conta_contato)
                        else:
                            lista_conta_contato = DadosRetornoCSV(len('Conta Contrato Coletiva nº'), texto.find('Conta Contrato Coletiva nº'), texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO).replace(".", "")
                         
                    #     print('lista_conta_contato: ', lista_conta_contato)
                    #    # print('textoContaContrato: ', textoContaContrato)
                    #     print('Conta Contrato Coletiva no: ', len('Conta Contrato Coletiva'))
                    #     print('texto.find(Conta Contrato Coletiva no)', texto.find('Conta  Contrato Coletiva'))
                    #     print('texto.find(A Iluminação Pública é de responsabilidade da Prefeitura)', texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'))

                        lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto.find('MÊS/ANO'), texto.find('VENCIMENTO')+1, TEXTO_COMPLETO)

                        lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto.find('TOTAL A PAGAR R$'), texto.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                        # writer.writerow([
                        #     lista_dados_cliente,
                        #     lista_end_unid_consum,
                        #     lista_num_nota_fiscal,
                        #     lista_num_Instalacao,
                        #     lista_classificacao,
                        #     lista_desc_nota_fiscal_gerar,
                        #     lista_desc_tarifa_gerar,
                        #     lista_inform_tributos_list_ICMS[0],
                        #     lista_inform_tributos_list_ICMS[1],
                        #     lista_inform_tributos_list_ICMS[2],
                        #     lista_inform_tributos_list_PIS[0],
                        #     lista_inform_tributos_list_PIS[1],
                        #     lista_inform_tributos_list_PIS[2],
                        #     lista_inform_tributos_list_COFINS[0],
                        #     lista_inform_tributos_list_COFINS[1],
                        #     lista_inform_tributos_list_COFINS[2],
                        #     lista_num_medidor,
                        #     lista_conta_contato,
                        #     lista_mes_ano,
                        #     lista_total_pagar
                        # ])

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
            
               if page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and \
                  page.extract_text().find('AVISO IMPORTANTE!') == -1 and \
                  page.extract_text().find('2 /   3') == -1 and \
                  page.extract_text().find('3 /   3') == -1:
         
                  TEXTO_COMPLETO = page.extract_text()      

                  try:                                         

                    # Dados do cliente
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_dados_cliente = DadosRetornoCSV(len('DADOS DO CLIENTE'), page.extract_text().find('DADOS DO CLIENTE'), page.extract_text().find('DATA DE VENCIMENTO'), TEXTO_COMPLETO)            
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_dados_cliente = DadosRetornoCSV(len('DADOS DO CLIENTE'), page.extract_text().find('DADOS DO CLIENTE'), page.extract_text().find('ENDEREÇO'), TEXTO_COMPLETO)

                    # print('lista_dados_cliente: ', lista_dados_cliente)

                    # Endereço Unidade Consumidora
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO DA UNIDADE CONSUMIDORA'), page.extract_text().find('ENDEREÇO DA UNIDADE CONSUMIDORA'), page.extract_text().find('RESERVADO AO FISCO'), TEXTO_COMPLETO)                      
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO'), page.extract_text().find('ENDEREÇO'), page.extract_text().find('DATA DE VENCIMENTO')+1, TEXTO_COMPLETO) 

                    # Número da Nota Fiscal
                    lista_num_nota_fiscal = DadosRetornoCSV(len('NÚMERO DA NOTA FISCAL'), page.extract_text().find('NÚMERO DA NOTA FISCAL'), page.extract_text().find('CONTA CONTRATO'), TEXTO_COMPLETO)          
                
                    # Nº da Instlação
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_num_Instalacao = DadosRetornoCSV(len('Nº DA INSTALAÇÃO'), page.extract_text().find('Nº DA INSTALAÇÃO'), page.extract_text().find('CLASSIFICAÇÃO'), TEXTO_COMPLETO)                      
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_num_Instalacao = DadosRetornoCSV(len('Nº DA INSTALAÇÃO'), page.extract_text().find('Nº DA INSTALAÇÃO'), page.extract_text().find('RESERVADO AO FISCO'), TEXTO_COMPLETO)
                    
                    # Classificação
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO'), page.extract_text().find('CLASSIFICAÇÃO'), page.extract_text().find('ENDEREÇO DA UNIDADE CONSUMIDORA'), TEXTO_COMPLETO)          
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO'), page.extract_text().find('CLASSIFICAÇÃO'), page.extract_text().find('DESCRIÇÃO DA NOTA FISCAL E INFORMAÇÕES IMPORTANTES'), TEXTO_COMPLETO)          

                    # Descrição da Nota Fiscal
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('DESCRIÇÃO DA NOTA FISCAL'), TEXTO_COMPLETO)
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('DESCRIÇÃO DA NOTA FISCAL E INFORMAÇÕES IMPORTANTES'), TEXTO_COMPLETO)
                    elif page.extract_text().find('REF:MÊS/ANO') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
                                                                                                                    
                    lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split())

                    # Tarifas Aplicadas
                    if page.extract_text().find('DATA PREVISTA DA PRÓXIMA LEITURA:') > 0:
                        lista_tarifas_aplicadas_tratada = DadosRetornoCSV(len('DATA PREVISTA DA PRÓXIMA LEITURA:')+11, page.extract_text().find('DATA PREVISTA DA PRÓXIMA LEITURA:'), page.extract_text().find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('NOTA FISCAL | FATURA | CONTA DE ENERGIA ELÉTRICA'), TEXTO_COMPLETO)
                    else:
                        lista_tarifas_aplicadas = DadosRetornoCSV(len('AJUSTECONSUMO'), page.extract_text().find('AJUSTECONSUMO'), page.extract_text().find('Tarifas Aplicadas'), TEXTO_COMPLETO)          
                        lista_tarifas_aplicadas_tratada = lista_tarifas_aplicadas.replace('(kWh)','').strip()

                    # Informações de Tributos
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_inform_tributos = DadosRetornoCSV(len('INFORMAÇÕES DE TRIBUTOS'), page.extract_text().find('INFORMAÇÕES DE TRIBUTOS'), page.extract_text().find('AUTENTICAÇÃO MECÂNICA'), TEXTO_COMPLETO)
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_inform_tributos = DadosRetornoCSV(len('CÁLCULO % IMPOSTO PIS/COFINS % IMPOSTO % IMPOSTO'), page.extract_text().find('CÁLCULO % IMPOSTO PIS/COFINS % IMPOSTO % IMPOSTO'), page.extract_text().find('1 /   3'), TEXTO_COMPLETO)
                    
                    lista_inform_tributos_tratado = " ".join(lista_inform_tributos.split()) # Retrar os espaços entre as palavras      
                    lista_inform_tributos_list = lista_inform_tributos_tratado.split(" ") # Converter em lista  

                    diferencaTotal = -1
                    lista_inform_tributos_list_aux = [' ',' ',' ',' ',' ',' ',' ',' ',' ']

                    for x in range(len(lista_inform_tributos_list)):                               
                        lista_inform_tributos_list_aux[diferencaTotal] = lista_inform_tributos_list[diferencaTotal]
                        diferencaTotal -=1                      

                    # Número do Medidor
                    if page.extract_text().find('CAT') > 0 and page.extract_text().find('AJUSTECONSUMO') > 0:
                        lista_num_medidor = DadosRetornoCSV(len('AJUSTECONSUMO'), page.extract_text().find('AJUSTECONSUMO'), page.extract_text().find('CAT'), TEXTO_COMPLETO)
                        lista_num_medidor_tratado = lista_num_medidor.replace('(kWh)','').strip()
                    else:
                        lista_num_medidor_tratado = ""

                    # Conta Contrato
                    lista_conta_contato = DadosRetornoCSV(len('CONTA CONTRATO'), page.extract_text().find('CONTA CONTRATO'), page.extract_text().find('Nº DO CLIENTE'), TEXTO_COMPLETO)
                
                    # Mês Ano
                    if page.extract_text().find('AUTENTICAÇÃO MECÂNICA') > 0:
                        lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), page.extract_text().find('MÊS/ANO'), page.extract_text().find('TOTAL A PAGAR(R$)')+1, TEXTO_COMPLETO) 
                    elif page.extract_text().find('1 /   3') > 0:
                        lista_mes_ano = DadosRetornoCSV(len('DATA DA EMISSÃO DA NOTA FISCAL'), page.extract_text().find('DATA DA EMISSÃO DA NOTA FISCAL')+4, page.extract_text().find('DATA DA APRESENTAÇÃO')+1, TEXTO_COMPLETO) 
                    else:
                        lista_mes_ano = DadosRetornoCSV(len('REF:MÊS/ANO'), pageAux.extract_text().find('REF:MÊS/ANO')+3, pageAux.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO_AUX) 

                    # Total a pagar
                    lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR (R$)'), page.extract_text().find('TOTAL A PAGAR (R$)'), page.extract_text().find('DATA DA EMISSÃO DA NOTA FISCAL'), TEXTO_COMPLETO)

                    # writer.writerow([
                    #         lista_dados_cliente, 
                    #         lista_end_unid_consum, 
                    #         lista_num_nota_fiscal, 
                    #         lista_num_Instalacao, 
                    #         lista_classificacao, 
                    #         lista_desc_nota_fiscal_tratado, 
                    #         lista_tarifas_aplicadas_tratada, 
                    #         lista_inform_tributos_list_aux[0], 
                    #         lista_inform_tributos_list_aux[1],
                    #         lista_inform_tributos_list_aux[2],
                    #         lista_inform_tributos_list_aux[3],
                    #         lista_inform_tributos_list_aux[4],
                    #         lista_inform_tributos_list_aux[5],
                    #         lista_inform_tributos_list_aux[6],
                    #         lista_inform_tributos_list_aux[7],
                    #         lista_inform_tributos_list_aux[8],
                    #         lista_num_medidor_tratado,
                    #         lista_conta_contato,
                    #         lista_mes_ano,
                    #         lista_total_pagar
                    # ])                   

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

    # st.download_button(
    #     label="📥 Baixar CSV",
    #     data=csv_bytes,
    #     file_name="relatorio_coelba.csv",
    #     mime="text/csv; charset=utf-8"
    # )

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
            # if col_name.lower() == "total a pagar" or \
            #                     "ICMS Base de Cálculo" or\
            #                     "ICMS Base 2" or \
            #                     "ICMS Base 3" or \
            #                     "ICMS Base 4" or \
            #                     "ICMS Base 5" or \
            #                     "ICMS Base 6" or \
            #                     "ICMS Base 7" or \
            #                     "ICMS Base 8" or \
            #                     "ICMS Base 9":
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
    #################################################