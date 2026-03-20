import streamlit as st
import pandas as pd
import io
import pdfplumber
from funcoes import DadosRetornoCSV
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURAÇÕES DA API ---
SCOPES = ['https://www.googleapis.com/auth/drive']
PASTA_CSV_ID = '1r6DtkBXm7TZyBOKnwE4tG_-j02N5V5VX'                
PASTA_XLSX_ID = '13ifqsQjGl2_M-VoOxTMtJI0JvIsDrwhy'

def autenticar_drive():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)

def upload_para_drive(service, nome_arquivo, conteudo, pasta_id, mimetype):
    try:
        file_metadata = {
            'name': nome_arquivo,
            'parents': [pasta_id]
        }

        media = MediaIoBaseUpload(
            conteudo, 
            mimetype=mimetype, 
            resumable=True
        )

        # O parâmetro entra aqui, logo após o media_body
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True,
            keepRevisionForever=False  # <--- INSERIDO AQUI
        ).execute()

        return file.get('id')
    except Exception as e:
        st.error(f"Erro no upload: {e}")
        return None

# --- SUA LÓGICA DE EXTRAÇÃO ADAPTADA ---
def DadosRetornoCSV(tamanho_label, pos_inicio, pos_fim, texto_completo):
    """Simulação da sua função original de extração de substrings"""
    if pos_inicio == -1 or pos_fim == -1:
        return ""
    resultado = texto_completo[pos_inicio + tamanho_label : pos_fim].strip()
    return resultado

def processar_pdfs(uploaded_files):
    dados_finais = []
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    # FUNÇÃO AUXILIAR DE SEGURANÇA
    def get_val(lista, index):
        """Retorna o valor da lista no índice informado ou vazio se não existir"""
        try:
            return lista[index]
        except IndexError:
            return ""

    for idx, pdf_file in enumerate(uploaded_files):
        with pdfplumber.open(pdf_file) as pdf:
            controlarPag = 0
            
            for page in pdf.pages:
                texto_da_pagina = page.extract_text()
                if not texto_da_pagina: continue

                if (texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    texto_da_pagina.find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    texto_da_pagina.find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    texto_da_pagina.find('DIC, FIC, DMIC e DICRI') == -1):
                    
                    TEXTO_COMPLETO = texto_da_pagina

                    # ... (suas extrações de cliente, endereço, etc permanecem iguais)
                    lista_dados_cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), texto_da_pagina.find('NOME DO CLIENTE:'), texto_da_pagina.find('ENDEREÇO:'), TEXTO_COMPLETO)
                    lista_end_unid_consum = DadosRetornoCSV(len('ENDEREÇO:'), texto_da_pagina.find('ENDEREÇO:'), texto_da_pagina.find('CÓDIGO DA')+1, TEXTO_COMPLETO)
                    lista_num_nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), texto_da_pagina.find('NOTA FISCAL N°'), texto_da_pagina.find('- SÉRIE'), TEXTO_COMPLETO)
                    lista_num_Instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), texto_da_pagina.find('INSTALAÇÃO'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)
                    lista_classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), texto_da_pagina.find('CLASSIFICAÇÃO:'), texto_da_pagina.find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)

                    # TRATAMENTO DA DESCRIÇÃO E TARIFAS (PONTO DE ERRO)
                    if texto_da_pagina.find('neoenergiacoelba.com.br') > 0:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto_da_pagina.find('neoenergiacoelba.com.br'), TEXTO_COMPLETO)
                    else:
                        lista_desc_nota_fiscal = DadosRetornoCSV(0, 0, texto_da_pagina.find('www.neoenergia.com'), TEXTO_COMPLETO)

                    lista_desc_nota_fiscal_tratado = " ".join(lista_desc_nota_fiscal.split())
                    lista_desc_nota_fiscal_tratado_list = lista_desc_nota_fiscal_tratado.split(" ")
                    
                    # Monta a string de tarifas com segurança
                    tarifas = [
                        get_val(lista_desc_nota_fiscal_tratado_list, 0),
                        get_val(lista_desc_nota_fiscal_tratado_list, 9),
                        get_val(lista_desc_nota_fiscal_tratado_list, 10),
                        get_val(lista_desc_nota_fiscal_tratado_list, 19)
                    ]
                    lista_desc_tarifa_gerar = " ".join(tarifas)

                    # TRIBUTOS (PONTO DE ERRO)
                    # ICMS
                    raw_icms = DadosRetornoCSV(len('ICMS'), texto_da_pagina.find('ICMS'), texto_da_pagina.find('CONSUMO / kWh'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_ICMS = " ".join(raw_icms.split()).split(" ")
                    
                    # PIS
                    raw_pis = DadosRetornoCSV(len('(%) PIS'), texto_da_pagina.find('(%) PIS'), texto_da_pagina.find('COFINS'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_PIS = " ".join(raw_pis.split()).split(" ")

                    # COFINS
                    raw_cofins = DadosRetornoCSV(len('COFINS'), texto_da_pagina.find('COFINS'), texto_da_pagina.find('ICMS'), TEXTO_COMPLETO)
                    lista_inform_tributos_list_COFINS = " ".join(raw_cofins.split()).split(" ")

                    # ... (medidor, mes_ano, total_pagar)
                    lista_num_medidor_tratado = DadosRetornoCSV(len('MEDIDOR kWh'), texto_da_pagina.find('MEDIDOR kWh'), texto_da_pagina.find('Energia Ativa'), TEXTO_COMPLETO) if texto_da_pagina.find('MEDIDOR kWh') > 0 else ""
                    lista_conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), texto_da_pagina.find('CÓDIGO DO CLIENTE'), texto_da_pagina.find(' DATAS DE LEITURAS'), TEXTO_COMPLETO).replace(".", "")
                    lista_mes_ano = DadosRetornoCSV(len('MÊS/ANO'), texto_da_pagina.find('MÊS/ANO'), texto_da_pagina.find('VENCIMENTO')+1, TEXTO_COMPLETO) 
                    lista_total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), texto_da_pagina.find('TOTAL A PAGAR R$'), texto_da_pagina.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                    # ADICIONANDO AO DICIONÁRIO COM get_val (SEGURANÇA TOTAL)
                    dados_finais.append({
                        "Conta Contrato": lista_conta_contato,
                        "Mês Ano": lista_mes_ano,
                        "Dados do cliente": lista_dados_cliente,
                        "Endereço da Unidade Consumidora": lista_end_unid_consum,
                        "Número da Nota Fiscal": lista_num_nota_fiscal,
                        "Número da Instalação": lista_num_Instalacao,
                        "Classificação": lista_classificacao,
                        "Descrição da Nota Fiscal": " ".join(lista_desc_nota_fiscal_tratado_list),
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
                        "Número do medidor": lista_num_medidor_tratado,
                        "Total a pagar": lista_total_pagar                                      
                    })                                                                                                                                                                                                                                                                                             
                    controlarPag += 1
                elif texto_da_pagina.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
                    controlarPag = 0
        
        progress_bar.progress((idx + 1) / total_files)
    
    return pd.DataFrame(dados_finais)

# --- INTERFACE PRINCIPAL ---
st.set_page_config(page_title="Conversor Neoenergia", layout="wide")
st.title("⚡ Extrator Neoenergia para Google Drive")

arquivos = st.file_uploader("Selecione os arquivos PDF das contas", type="pdf", accept_multiple_files=True)

if arquivos:
    if st.button("Iniciar Processamento e Upload"):
        try:
            drive_service = autenticar_drive()
            
            # Executa sua lógica
            df_resultado = processar_pdfs(arquivos)
            
            if not df_resultado.empty:
                st.write("### Prévia dos Dados Extraídos", df_resultado.head())

                # 1. Gerar e Enviar CSV
                csv_buffer = io.BytesIO()
                df_resultado.to_csv(csv_buffer, index=False, encoding='utf-8')
                csv_buffer.seek(0)
                id_csv = upload_para_drive(drive_service, "contas_processadas.csv", csv_buffer, PASTA_CSV_ID, 'text/csv')
                st.success(f"✅ CSV enviado com sucesso! (ID: {id_csv})")

                # 2. Gerar e Enviar XLSX
                xlsx_buffer = io.BytesIO()
                with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False)
                xlsx_buffer.seek(0)
                id_xlsx = upload_para_drive(drive_service, "contas_finais.xlsx", xlsx_buffer, PASTA_XLSX_ID, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                st.success(f"✅ Excel enviado com sucesso! (ID: {id_xlsx})")

                # 3. Download Local
                xlsx_buffer.seek(0)
                st.download_button(
                    label="📥 Baixar Excel Agora",
                    data=xlsx_buffer,
                    file_name="contas_neoenergia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nenhum dado foi extraído com os filtros aplicados.")
                
        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")