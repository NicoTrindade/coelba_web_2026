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
                if not page.extract_text(): continue

                # Implementação da sua lógica de filtros
                if (page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 0 and
                    page.extract_text().find('Fale com a gente! | Nossos Canais de Atendimento') == -1 and
                    page.extract_text().find('TELEATENDIMENTO: Emergencial 116') == -1 and
                    page.extract_text().find('DIC, FIC, DMIC e DICRI') == -1):
                                
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

                    # Organizando na lista de resultados (Equivalente ao seu writer.writerow)
                    dados_finais.append({
                        "Conta Contrato": lista_conta_contato,
                        "Mês Ano": lista_mes_ano,
                        "Dados do cliente": lista_dados_cliente,
                        "Endereço da Unidade Consumidora": lista_end_unid_consum,
                        "Número da Nota Fiscal": lista_num_nota_fiscal,
                        "Número da Instalação": lista_num_Instalacao,
                        "Classificação": lista_classificacao,
                        "Descrição da Nota Fiscal": lista_desc_nota_fiscal_gerar,
                        "Tarifas Aplicadas": lista_desc_tarifa_gerar,
                        "ICMS Base de Cálculo": lista_inform_tributos_list_ICMS[0] if len(lista_inform_tributos_list_ICMS[0])>0 else "",
                        "ICMS Base 2": lista_inform_tributos_list_ICMS[1] if len(lista_inform_tributos_list_ICMS[1])>0 else "",
                        "ICMS Base 3": lista_inform_tributos_list_ICMS[2] if len(lista_inform_tributos_list_ICMS[2])>0 else "",
                        "ICMS Base 4": lista_inform_tributos_list_PIS[0] if len(lista_inform_tributos_list_PIS[0])>0 else "",
                        "ICMS Base 5": lista_inform_tributos_list_PIS[1] if len(lista_inform_tributos_list_PIS[1])>0 else "",
                        "ICMS Base 6": lista_inform_tributos_list_PIS[2] if len(lista_inform_tributos_list_PIS[2])>0 else "",
                        "ICMS Base 7": lista_inform_tributos_list_COFINS[0] if len(lista_inform_tributos_list_COFINS[0])>0 else "",
                        "ICMS Base 8": lista_inform_tributos_list_COFINS[1] if len(lista_inform_tributos_list_COFINS[1])>0 else "",
                        "ICMS Base 9": lista_inform_tributos_list_COFINS[2] if len(lista_inform_tributos_list_COFINS[2])>0 else "",
                        "Número do medidor": lista_num_medidor_tratado,
                        "Total a pagar": lista_total_pagar                     
                    })                                                                                                                                                                                                        
                    controlarPag += 1
                elif page.extract_text().find('DOCUMENTO PARA PAGAMENTO DA CONTA COLETIVA') == -1 and controlarPag == 1:
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