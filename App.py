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
    
    # Barra de progresso do Streamlit
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    for idx, pdf_file in enumerate(uploaded_files):
        with pdfplumber.open(pdf_file) as pdf:
            controlarPag = 0 # Sua variável de controle
            
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue

                # Implementação da sua lógica de filtros
                if (text.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    text.find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    text.find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    text.find('DIC, FIC, DMIC e DICRI') == -1):
                    
                    TEXTO_COMPLETO = text

                    # Extrações conforme seu código
                    cliente = DadosRetornoCSV(len('NOME DO CLIENTE:'), text.find('NOME DO CLIENTE:'), text.find('ENDEREÇO:'), TEXTO_COMPLETO)
                    endereco = DadosRetornoCSV(len('ENDEREÇO:'), text.find('ENDEREÇO:'), text.find('CÓDIGO DA')+1, TEXTO_COMPLETO)
                    nota_fiscal = DadosRetornoCSV(len('NOTA FISCAL N°'), text.find('NOTA FISCAL N°'), text.find('- SÉRIE'), TEXTO_COMPLETO)
                    instalacao = DadosRetornoCSV(len('INSTALAÇÃO'), text.find('INSTALAÇÃO'), text.find('CÓDIGO DO CLIENTE'), TEXTO_COMPLETO)
                    classificacao = DadosRetornoCSV(len('CLASSIFICAÇÃO:'), text.find('CLASSIFICAÇÃO:'), text.find('TIPO DE FORNECIMENTO:'), TEXTO_COMPLETO)

                    # Descrição da Nota
                    pos_web = text.find('neoenergiacoelba.com.br') if text.find('neoenergiacoelba.com.br') > 0 else text.find('www.neoenergia.com')
                    desc_nota = DadosRetornoCSV(0, 0, pos_web, TEXTO_COMPLETO)
                    desc_lista = " ".join(desc_nota.split()).split(" ")
                    
                    # Tributos (ICMS, PIS, COFINS)
                    icms_raw = DadosRetornoCSV(len('ICMS'), text.find('ICMS'), text.find('CONSUMO / kWh'), TEXTO_COMPLETO)
                    icms_list = " ".join(icms_raw.split()).split(" ")
                    
                    pis_raw = DadosRetornoCSV(len('(%) PIS'), text.find('(%) PIS'), text.find('COFINS'), TEXTO_COMPLETO)
                    pis_list = " ".join(pis_raw.split()).split(" ")

                    cofins_raw = DadosRetornoCSV(len('COFINS'), text.find('COFINS'), text.find('ICMS'), TEXTO_COMPLETO)
                    cofins_list = " ".join(cofins_raw.split()).split(" ")

                    # Medidor, Mes/Ano e Total
                    medidor = DadosRetornoCSV(len('MEDIDOR kWh'), text.find('MEDIDOR kWh'), text.find('Energia Ativa'), TEXTO_COMPLETO) if text.find('MEDIDOR kWh') > 0 else ""
                    conta_contato = DadosRetornoCSV(len('CÓDIGO DO CLIENTE'), text.find('CÓDIGO DO CLIENTE'), text.find(' DATAS DE LEITURAS'), TEXTO_COMPLETO).replace(".", "")
                    mes_ano = DadosRetornoCSV(len('MÊS/ANO'), text.find('MÊS/ANO'), text.find('VENCIMENTO')+1, TEXTO_COMPLETO)
                    total_pagar = DadosRetornoCSV(len('TOTAL A PAGAR R$'), text.find('TOTAL A PAGAR R$'), text.find('Cadastra-se e receba'), TEXTO_COMPLETO)

                    # Organizando na lista de resultados (Equivalente ao seu writer.writerow)
                    dados_finais.append({
                        "Código Cliente": conta_contato,
                        "Mês/Ano": mes_ano,
                        "Cliente": cliente,
                        "Endereço": endereco,
                        "Nota Fiscal": nota_fiscal,
                        "Instalação": instalacao,
                        "Classificação": classificacao,
                        "ICMS_0": icms_list[0] if len(icms_list)>0 else "",
                        "PIS_0": pis_list[0] if len(pis_list)>0 else "",
                        "COFINS_0": cofins_list[0] if len(cofins_list)>0 else "",
                        "Medidor": medidor,
                        "Total a Pagar": total_pagar
                    })
                    
                    controlarPag += 1
                elif text.find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
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