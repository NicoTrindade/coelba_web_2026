import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- FUNÇÕES DE APOIO (Preservando sua lógica original) ---

def DadosRetornoCSV(tamanho_label, inicio, fim, texto):
    if inicio == -1 or fim == -1:
        return ""
    return texto[inicio + tamanho_label:fim].strip()

def get_val(lista, index):
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

                # Filtros de segurança que você definiu
                if (texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    texto_da_pagina.find('Fale com a gente!') == -1 and
                    texto_da_pagina.find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    texto_da_pagina.find('DIC, FIC, DMIC e DICRI') == -1):
                    
                    TEXTO_COMPLETO = texto_da_pagina

                    # 1. Extrações Básicas
                    lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), texto_da_pagina.find('NOME DO CLIENTE:'), texto_da_pagina.find('ENDEREÇO:'), TEXTO_COMPLETO)
                    lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), texto_da_pagina.find('ENDEREÇO:'), texto_da_pagina.find('CÓDIGO DA')+1, TEXTO_COMPLETO)
                    lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), texto_da_pagina.find('NOTA FISCAL N°'), texto_da_pagina.find('- SÉRIE'), TEXTO_COMPLETO)
                    lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), texto_da_pagina.find('INSTALAÇÃO'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)
                    lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), texto_da_pagina.find('CLASSIFICAÇÃO:'), texto_da_pagina.find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)

                    # 2. Descrição e Tarifas
                    link_pos = texto_da_pagina.find('neoenergiacoelba.com.br') if texto_da_pagina.find('neoenergiacoelba.com.br') > 0 else texto_da_pagina.find('www.neoenergia.com')
                    lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, link_pos, TEXTO_COMPLETO)
                    desc_tratado_list = " ".join(lista_desc_nota_fiscal.split()).split(" ")
                    
                    tarifas_str = " ".join([get_val(desc_tratado_list, 0), get_val(desc_tratado_list, 9), get_val(desc_tratado_list, 10), get_val(desc_tratado_list, 19)])

                    # 3. Tributos (Sua lógica de PIS/COFINS/ICMS)
                    raw_icms = DadosRetornoCSV(len('ICMS'), texto_da_pagina.find('ICMS'), texto_da_pagina.find('CONSUMO / kWh'), TEXTO_COMPLETO)
                    list_ICMS = " ".join(raw_icms.split()).split(" ")
                    
                    raw_pis = DadosRetornoCSV(len('(%) PIS'), texto_da_pagina.find('(%) PIS'), texto_da_pagina.find('COFINS'), TEXTO_COMPLETO)
                    list_PIS = " ".join(raw_pis.split()).split(" ")

                    raw_cofins = DadosRetornoCSV(len('COFINS'), texto_da_pagina.find('COFINS'), texto_da_pagina.find('ICMS'), TEXTO_COMPLETO)
                    list_COFINS = " ".join(raw_cofins.split()).split(" ")

                    # 4. Dados Complementares
                    medidor = DadosRetornoCSV(len('MEDIDOR kWh'), texto_da_pagina.find('MEDIDOR kWh'), texto_da_pagina.find('Energia Ativa'), TEXTO_COMPLETO) if texto_da_pagina.find('MEDIDOR kWh') > 0 else ""
                    conta_contrato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), texto_da_pagina.find(' DATAS DE LEITURAS'), TEXTO_COMPLETO).replace(".", "")
                    mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto_da_pagina.find('MÊS/ANO'), texto_da_pagina.find('VENCIMENTO')+1, TEXTO_COMPLETO) 
                    total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto_da_pagina.find('TOTAL A PAGAR R$'), texto_da_pagina.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                    # Montagem do Dicionário com TODOS os seus campos
                    dados_finais.append({
                        "Conta Contrato": conta_contrato,
                        "Mês Ano": mes_ano,
                        "Dados do cliente": lista_dados_cliente,
                        "Endereço": lista_end_unid_consum,
                        "NF": lista_num_nota_fiscal,
                        "Instalação": lista_num_Instalacao,
                        "Classificação": lista_classificacao,
                        "Descrição NF": " ".join(desc_tratado_list),
                        "Tarifas Aplicadas": tarifas_str,
                        "ICMS Base 1": get_val(list_ICMS, 0),
                        "ICMS Base 2": get_val(list_ICMS, 1),
                        "ICMS Base 3": get_val(list_ICMS, 2),
                        "ICMS Base 4": get_val(list_PIS, 0),
                        "ICMS Base 5": get_val(list_PIS, 1),
                        "ICMS Base 6": get_val(list_PIS, 2),
                        "ICMS Base 7": get_val(list_COFINS, 0),
                        "ICMS Base 8": get_val(list_COFINS, 1),
                        "ICMS Base 9": get_val(list_COFINS, 2),
                        "Medidor": medidor,
                        "Total a Pagar": total_pagar
                    })
                    controlarPag += 1
                elif texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
                    controlarPag = 0
        
        progress_bar.progress((idx + 1) / total_files)
    
    return pd.DataFrame(dados_finais)

# --- INTERFACE PRINCIPAL ---
st.set_page_config(page_title="Conversor Neoenergia", layout="wide")
st.title("⚡ Extrator Neoenergia Coelba (Oficial)")

arquivos = st.file_uploader("Selecione os PDFs das contas", type="pdf", accept_multiple_files=True)

if arquivos:
    if st.button("🚀 Iniciar Processamento"):
        df_resultado = processar_pdfs(arquivos)
        
        if not df_resultado.empty:
            st.write("### Prévia dos Dados Extraídos")
            st.dataframe(df_resultado)

            # --- GERAÇÃO DO ARQUIVO XLSX PARA DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resultado.to_excel(writer, index=False, sheet_name='Contas')
            
            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=output.getvalue(),
                file_name="contas_coelba_consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado extraído. Verifique o formato dos PDFs.")