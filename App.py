import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- FUNÇÕES DE APOIO ---

def DadosRetornoCSV(tamanho_label, inicio, fim, texto):
    """Extrai texto entre dois pontos com base no tamanho da etiqueta inicial"""
    if inicio == -1 or fim == -1:
        return ""
    resultado = texto[inicio + tamanho_label:fim].strip()
    return " ".join(resultado.split())  # Limpa espaços e quebras de linha

def get_val(lista, index):
    """Retorna o valor da lista no índice ou vazio se não existir"""
    try:
        return lista[index] if index < len(lista) else ""
    except:
        return ""

def processar_pdfs(uploaded_files):
    dados_finais = []
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    for idx, pdf_file in enumerate(uploaded_files):
        with pdfplumber.open(pdf_file) as pdf:
            controlarPag = 0
            
            for page in pdf.pages:
                texto_da_pagina = page.extract_text()
                if not texto_da_pagina: continue

                # Sua lógica de filtragem original
                if (texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    texto_da_pagina.find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    texto_da_pagina.find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    texto_da_pagina.find('DIC, FIC, DMIC e DICRI') == -1):
                    
                    TEXTO_COMPLETO = texto_da_pagina

                    # Extrações Principais
                    lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), texto_da_pagina.find('NOME DO CLIENTE:'), texto_da_pagina.find('ENDEREÇO:'), TEXTO_COMPLETO)
                    lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), texto_da_pagina.find('ENDEREÇO:'), texto_da_pagina.find('CÓDIGO DA')+1, TEXTO_COMPLETO)
                    lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), texto_da_pagina.find('NOTA FISCAL N°'), texto_da_pagina.find('- SÉRIE'), TEXTO_COMPLETO)
                    lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), texto_da_pagina.find('INSTALAÇÃO'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)
                    lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), texto_da_pagina.find('CLASSIFICAÇÃO:'), texto_da_pagina.find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)

                    # Tratamento de Descrição e Tarifas
                    link_ref = 'neoenergiacoelba.com.br' if texto_da_pagina.find('neoenergiacoelba.com.br') > 0 else 'www.neoenergia.com'
                    lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto_da_pagina.find(link_ref), TEXTO_COMPLETO)
                    lista_desc_nota_fiscal_tratado_list = " ".join(lista_desc_nota_fiscal.split()).split(" ")
                    
                    tarifas = [
                        get_val(lista_desc_nota_fiscal_tratado_list, 0),
                        get_val(lista_desc_nota_fiscal_tratado_list, 9),
                        get_val(lista_desc_nota_fiscal_tratado_list, 10),
                        get_val(lista_desc_nota_fiscal_tratado_list, 19)
                    ]
                    lista_desc_tarifa_gerar = " ".join(tarifas)

                    # Tributos
                    raw_icms = DadosRetornoCSV(len('ICMS'), texto_da_pagina.find('ICMS'), texto_da_pagina.find('CONSUMO / kWh'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_ICMS = " ".join(raw_icms.split()).split(" ")
                    
                    raw_pis = DadosRetornoCSV(len('(%) PIS'), texto_da_pagina.find('(%) PIS'), texto_da_pagina.find('COFINS'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_PIS = " ".join(raw_pis.split()).split(" ")

                    raw_cofins = DadosRetornoCSV(len('COFINS'), texto_da_pagina.find('COFINS'), texto_da_pagina.find('ICMS'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_COFINS = " ".join(raw_cofins.split()).split(" ")

                    # Dados de Rodapé/Conta
                    lista_num_medidor_tratado = DadosRetornoCSV(len('MEDIDOR kWh'), texto_da_pagina.find('MEDIDOR kWh'), texto_da_pagina.find('Energia Ativa'), TEXTO_COMPLETO) if texto_da_pagina.find('MEDIDOR kWh') > 0 else ""
                    lista_conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), texto_da_pagina.find(' DATAS DE LEITURAS'), TEXTO_COMPLETO).replace(".", "")
                    lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto_da_pagina.find('MÊS/ANO'), texto_da_pagina.find('VENCIMENTO')+1, TEXTO_COMPLETO) 
                    lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto_da_pagina.find('TOTAL A PAGAR R$'), texto_da_pagina.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                    # Montagem do Registro
                    dados_finais.append({
                        "Conta Contrato": lista_conta_contato,
                        "Mês Ano": lista_mes_ano,
                        "Dados do cliente": lista_dados_cliente,
                        "Endereço Unid. Consumidora": lista_end_unid_consum,
                        "Nota Fiscal": lista_num_nota_fiscal,
                        "Instalação": lista_num_Instalacao,
                        "Classificação": lista_classificacao,
                        "Descrição da NF": " ".join(lista_desc_nota_fiscal_tratado_list),
                        "Tarifas Aplicadas": lista_desc_tarifa_gerar,
                        "ICMS Base de Cálculo": get_val(lista_inform_tributos_list_ICMS, 0),
                        "ICMS Base 2": get_val(lista_inform_tributos_list_ICMS, 1),
                        "ICMS Base 3": get_val(lista_inform_tributos_list_ICMS, 2),
                        "ICMS Base 4": get_val(lista_inform_tributos_list_PIS, 0),
                        "ICMS Base 5": get_val(lista_inform_tributos_list_PIS, 1),
                        "ICMS Base 6": get_val(lista_inform_tributos_list_PIS, 2),
                        "ICMS Base 7": get_val(lista_inform_tributos_list_COFINS, 0),
                        "ICMS Base 8": get_val(lista_inform_tributos_list_COFINS, 1),
                        "ICMS Base 9": get_val(lista_inform_tributos_list_COFINS, 2),
                        "Número do Medidor": lista_num_medidor_tratado,
                        "Total a Pagar": lista_total_pagar                                     
                    })                                                                                                                                                                                                                                                                                                      
                    controlarPag += 1
                elif texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
                    controlarPag = 0
        
        progress_bar.progress((idx + 1) / total_files)
    
    return pd.DataFrame(dados_finais)

# --- INTERFACE STREAMLIT ---

st.set_page_config(page_title="Extrator Coelba", layout="wide")
st.title("⚡ Extrator Neoenergia Coelba para Excel")
st.markdown("Faça o upload dos seus arquivos PDF para consolidar os dados em uma única planilha.")

arquivos = st.file_uploader("Selecione os arquivos PDF", type="pdf", accept_multiple_files=True)

if arquivos:
    if st.button("🚀 Iniciar Extração"):
        df_resultado = processar_pdfs(arquivos)
        
        if not df_resultado.empty:
            st.success(f"Concluído! {len(df_resultado)} faturas processadas.")
            st.write("### Prévia dos Dados")
            st.dataframe(df_resultado.head())

            # Preparação do arquivo XLSX para download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_resultado.to_excel(writer, index=False, sheet_name='Faturas_Extraidas')
            
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="contas_coelba_extraidas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado foi extraído. Verifique se os arquivos seguem o padrão esperado.")
