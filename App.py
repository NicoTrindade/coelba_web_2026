from funcoes import DadosRetornoCSV
from utils.csv_to_excel import process_csv_files
from PyPDF2 import PdfReader
from io import StringIO
from google_drive import process_csv_and_excel
import streamlit as st
import pandas as pd
import csv
import os
import codecs

st.set_page_config(page_title="Extrator Inteligente de Faturas", layout="wide")

st.title("⚡ Extrator Inteligente de Faturas PDF")

st.write(
    "Envie um ou mais PDFs de faturas. "
    "A aplicação irá extrair automaticamente os dados estruturados."
)

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:

    results = []

    progress = st.progress(0)

    for i, file in enumerate(uploaded_files):
        st.write(f"📄 Processando arquivo: {file.name}")
     
        RELATORIO_COELBA = file.name    
            
        arquivoCSV = RELATORIO_COELBA.replace(".pdf","") + '.csv'

        CAMINHO_CSV = arquivoCSV          
              
        reader = PdfReader(file) 
        page = reader.pages

        totalRegistros = len(reader.pages)   

        st.write("📑 Boletos encontrados:", (totalRegistros-1)/2)   
         
        lista_cabecalho = ['Conta Contrato', 
                        'Mês Ano', 
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
                        'Total a pagar']

        #with open(CAMINHO_CSV, 'w', newline='') as csvfile:
        #   csv.DictWriter(csvfile, fieldnames=lista_cabecalho, quoting=csv.QUOTE_ALL, delimiter=',').writeheader()         
         
        contPag = 0
        controlarPag = 0
      
        output = StringIO()

        # escreve header na memória
        writer = csv.writer(output, quoting=csv.QUOTE_ALL, delimiter=',')                          
            
        for page in reader.pages:                
      
           if (page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
              page.extract_text().find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
              page.extract_text().find('TELEATENDIMENTO: Emergencial 116') == -1 and
              page.extract_text().find('DIC, FIC, DMIC e DICRI') == -1) :
                           
              TEXTO_COMPLETO = page.extract_text()   

              # Dados do cliente
              lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), page.extract_text().find('NOME DO CLIENTE:'), page.extract_text().find('ENDEREÇO:'), TEXTO_COMPLETO)            

              # Endereço Unidade Consumidora
              lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), page.extract_text().find('ENDEREÇO:'), page.extract_text().find('CÓDIGO DA')+1, TEXTO_COMPLETO)  # +1 por conta do caractér especial, pois, não está considerando ara contagem      

              # Número da Nota Fiscal
              lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), page.extract_text().find('NOTA FISCAL N°'), page.extract_text().find('- SÉRIE'), TEXTO_COMPLETO)          
         
              # Nº da Instlação
              lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), page.extract_text().find('INSTALAÇÃO'), page.extract_text().find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)                      
            
              # Classificação
              lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), page.extract_text().find('CLASSIFICAÇÃO:'), page.extract_text().find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)          

              # Descrição da Nota Fiscal
               
              if page.extract_text().find('neoenergiacoelba.com.br') > 0:
                 lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
              else:
                 lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, page.extract_text().find('www.neoenergia.com'), TEXTO_COMPLETO)

              lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split()) # Retrar os espaços entre as palavras
              lista_desc_nota_fiscal_tratado_list = lista_desc_nota_fiscal_tratado.split(" ") # Converter em lista                            
                     
              lista_desc_nota_fiscal_gerar = ""

              for lista_separados in lista_desc_nota_fiscal_tratado_list:
                 lista_desc_nota_fiscal_gerar = lista_desc_nota_fiscal_gerar + " " + lista_separados
            
              # Tarifas Aplicadas          
              lista_desc_tarifa_separados = []
              lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[0])
              lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[9])
              lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[10])
              lista_desc_tarifa_separados.append(lista_desc_nota_fiscal_tratado_list[19])
            
              lista_desc_tarifa_gerar = ""
              for lista_separados in lista_desc_tarifa_separados:
                 lista_desc_tarifa_gerar = lista_desc_tarifa_gerar + " " + lista_separados
            
              # Informações de Tributos            
              lista_inform_tributos_ICMS = DadosRetornoCSV(len('ICMS'), page.extract_text().find('ICMS'), page.extract_text().find('CONSUMO / kWh'), TEXTO_COMPLETO)          
              lista_inform_tributos_ICMS_tratado = " ".join(lista_inform_tributos_ICMS.split()) # Retrar os espaços entre as palavras      
              lista_inform_tributos_list_ICMS = lista_inform_tributos_ICMS_tratado.split(" ") # Converter em lista                            

              lista_inform_tributos_PIS = DadosRetornoCSV(len('(%) PIS'), page.extract_text().find('(%) PIS'), page.extract_text().find('COFINS'), TEXTO_COMPLETO)          
              lista_inform_tributos_PIS_tratado = " ".join(lista_inform_tributos_PIS.split()) # Retrar os espaços entre as palavras      
              lista_inform_tributos_list_PIS = lista_inform_tributos_PIS_tratado.split(" ") # Converter em lista   

              lista_inform_tributos_COFINS = DadosRetornoCSV(len('COFINS'), page.extract_text().find('COFINS'), page.extract_text().find('ICMS'), TEXTO_COMPLETO)          
              lista_inform_tributos_COFINS_tratado = " ".join(lista_inform_tributos_COFINS.split()) # Retrar os espaços entre as palavras      
              lista_inform_tributos_list_COFINS = lista_inform_tributos_COFINS_tratado.split(" ") # Converter em lista                         

              # Número do Medidor              
              if page.extract_text().find('MEDIDOR kWh') > 0 and page.extract_text().find('Energia Ativa') > 0:
                 lista_num_medidor_tratado = DadosRetornoCSV(len('MEDIDOR kWh'), page.extract_text().find('MEDIDOR kWh'), page.extract_text().find('Energia Ativa'), TEXTO_COMPLETO)
                 # lista_num_medidor_tratado = lista_num_medidor.replace('(kWh)','').strip()
              else:
                 lista_num_medidor_tratado = ""
              lista_conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), page.extract_text().find('CÓDIGO DO CLIENTE'), page.extract_text().find(' DATAS DE LEITURAS'), TEXTO_COMPLETO)

              lista_conta_contato = lista_conta_contato.replace(".","")
      
              # Mês Ano
              lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), page.extract_text().find('MÊS/ANO'), page.extract_text().find('VENCIMENTO')+1, TEXTO_COMPLETO) 
            
              # Total a pagar
              lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), page.extract_text().find('TOTAL A PAGAR R$'), page.extract_text().find('Cadastra-se e receba'), TEXTO_COMPLETO)                      
            
              # Implantação nova para google Drive
         
              """  writer = csv.writer(csvfile, quoting=csv.QUOTE_ALL, delimiter=',').writerow([lista_conta_contato,
                                                                                       lista_mes_ano,
                                                                                       lista_dados_cliente, 
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
                                                                                       lista_num_medidor_tratado,                                                                                 
                                                                                       lista_total_pagar])   """
              # Implantação nova para google Drive
              writer.writerow([lista_conta_contato,
                           lista_mes_ano,
                           lista_dados_cliente, 
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
                           lista_num_medidor_tratado,                                                                                 
                           lista_total_pagar])
            
               
              controlarPag += 1
              results.append("Ok")               
           elif page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:      
              controlarPag = 0           

           contPag += 1
           progress.progress((i + 1) / (totalRegistros-1)/2)

        csv_string = output.getvalue()
        csv_bytes = csv_string.encode("utf-8")

        CSV_FOLDER_ID = "1r6DtkBXm7TZyBOKnwE4tG_-j02N5V5VX"
        
        EXCEL_FOLDER_ID = "13ifqsQjGl2_M-VoOxTMtJI0JvIsDrwhy"

        csv_link, excel_link = process_csv_and_excel(
                                                   csv_bytes=csv_bytes,
                                                   file_name="relatorio.csv",
                                                   csv_folder=CSV_FOLDER_ID,
                                                   excel_folder=EXCEL_FOLDER_ID
                                                   )

        if results:
           st.write("✅ Total de faturas extraídas:", len(results))
           st.success("Extração finalizada!")          

           st.success("Arquivos enviados com sucesso!")
           st.markdown(f"📄 CSV: [Abrir]({csv_link})")
           st.markdown(f"📊 Excel: [Abrir]({excel_link})")
