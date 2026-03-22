import streamlit as st
from PyPDF2 import PdfReader
from io import StringIO
import csv

from funcoes import DadosRetornoCSV  # sua função existente

st.set_page_config(page_title="Extrator COELBA", layout="wide")

st.title("⚡ Extrator Inteligente de Faturas PDF")

uploaded_files = st.file_uploader(
    "Selecione os PDFs",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:

    #output = StringIO()
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

        for page in reader.pages:

            texto = page.extract_text()

            if texto is None:
                continue

            if texto.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0:

                TEXTO_COMPLETO = texto

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

                    lista_conta_contato = DadosRetornoCSV(len('Conta Contrato Coletiva nº'), texto.find('Conta Contrato Coletiva nº'), texto.find('A Iluminação Pública é de responsabilidade da Prefeitura'), TEXTO_COMPLETO).replace(".", "")

                    lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto.find('MÊS/ANO'), texto.find('VENCIMENTO')+1, TEXTO_COMPLETO)

                    lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto.find('TOTAL A PAGAR R$'), texto.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                    writer.writerow([
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

        progress.progress((i + 1) / total_files)

    # CSV final
    csv_string = output.getvalue()
    #csv_bytes = csv_string.encode("utf-8")
    csv_bytes = csv_string.encode("utf-8-sig")

    st.success(f"✅ Total de boletos extraídos: {total_boletos}")

    # 📥 BOTÃO DOWNLOAD
    # st.download_button(
    #     label="📥 Baixar CSV",
    #     data=csv_bytes,
    #     file_name="relatorio_coelba.csv",
    #     mime="text/csv"
    # )

    st.download_button(
        label="📥 Baixar CSV",
        data=csv_bytes,
        file_name="relatorio_coelba.csv",
        mime="text/csv; charset=utf-8"
    )