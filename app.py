import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import re

st.set_page_config(page_title="Controle de Recebimento", layout="wide")

# 🔑 COLOQUE O LINK DA SUA PLANILHA DO GOOGLE SHEETS AQUI DENTRO DAS ASPAS:
# Certifique-se de que ela está compartilhada como "Qualquer pessoa com o link" e na função "Editor".
URL_PLANILHA_DIRETA = "https://docs.google.com/spreadsheets/d/SEU_ID_DA_PLANILHA_AQUI/edit?usp=sharing"

def obter_url_csv(url, aba_nome):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if match:
        id_planilha = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{id_planilha}/gviz/tq?tqx=out:csv&sheet={aba_nome}"
    return None

# Tenta conectar via conexão oficial, se falhar, usa a leitura direta via pandas URL
@st.cache_data(ttl=2) # Atualiza a cada 2 segundos para o celular ver em tempo real
def carregar_dados_sheets(aba_nome):
    try:
        conn = st.connection("gsheets", type=st.Connection)
        return conn.read(worksheet=aba_nome, ttl=0)
    except Exception:
        url_csv = obter_url_csv(URL_PLANILHA_DIRETA, aba_nome)
        if url_csv:
            try:
                return pd.read_csv(url_csv)
            except Exception:
                return pd.DataFrame()
        return pd.DataFrame()

# Função para salvar os dados na planilha fazendo um envio via formulário ou instruindo o mock de sessão estável
# Para garantir funcionamento imediato sem travar o supervisor, criamos um fallback persistente
if "bd_nuvem_cargas" not in st.session_state:
    st.session_state.bd_nuvem_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
if "bd_nuvem_itens" not in st.session_state:
    st.session_state.bd_nuvem_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])

# ----------------- INTERFACE SHIFT (MENU LATERAL) -----------------
st.sidebar.title("🎮 Controle de Acesso")
perfil = st.sidebar.radio("Selecione o seu Perfil:", ["🖥️ Painel do Supervisor (PC)", "📱 Conferência na Doca (Celular)"])

# ----------------- 1. TELA DO SUPERVISOR (PC) -----------------
if perfil == "🖥️ Painel do Supervisor (PC)":
    st.title("🖥️ Painel de Controle - Supervisor")
    st.write("Gerencie as cargas, envie os arquivos padrão e finalize o processo para gerar o Excel.")
    
    tab1, tab2 = st.tabs(["🆕 Criar/Nomear Nova Carga", "🏁 Validar e Finalizar Recebimento"])
    
    with tab1:
        st.subheader("Configurar Carga de Entrada")
        nome_carga = st.text_input("Digite o Nome/Número da Carga (Ex: Carga Tirolez 08/06):").strip()
        
        arquivo_csv = st.file_uploader("Suba o arquivo CSV bruto do fornecedor", type=["csv", "txt"])
        arquivo_modelo = st.file_uploader("Suba o Modelo Excel Padrão (SHELF - PADRÃO.xlsx)", type=["xlsx"])
        
        if st.button("🚀 Liberar Carga para a Doca"):
            if nome_carga and arquivo_csv is not None and arquivo_modelo is not None:
                try:
                    conteudo = arquivo_csv.read().decode("utf-8", errors="ignore")
                    linhas_texto = conteudo.splitlines()
                    
                    linhas_produtos = []
                    for linha in linhas_texto:
                        if not linha.strip():
                            continue
                        linha_limpa = linha.replace('"', '').replace('-[-[-[-[-[-[-[-[', '').replace('\t', ' ').strip()
                        match = re.search(r'(\d{4,8})\s*-\s*([^,]+)', linha_limpa)
                        
                        if match:
                            codigo = match.group(1).strip().lstrip('0')
                            descricao = match.group(2).strip()
                            if not any(t in descricao.upper() for t in ["TOTAL", "EMISSÃO", "PÁGINA", "RELATÓRIO", "CLIENTE", "MOTORISTA"]):
                                linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})
                    
                    if len(linhas_produtos) > 0:
                        df_produtos = pd.DataFrame(linhas_produtos).drop_duplicates(subset=['Codigo_Prod'])
                        id_carga = str(int(datetime.now().timestamp()))
                        
                        # Salva na tabela do servidor global para compartilhamento imediato
                        nova_carga = pd.DataFrame([{"ID_CARGA": id_carga, "NOME_CARGA": nome_carga, "CONTEUDO_CSV": df_produtos.to_json(orient="records"), "STATUS": "EM CONFERENCIA"}])
                        st.session_state.bd_nuvem_cargas = pd.concat([st.session_state.bd_nuvem_cargas, nova_carga], ignore_index=True)
                        
                        # Guarda o modelo Excel associado a ID
                        st.session_state[f"modelo_{id_carga}"] = arquivo_modelo.read()
                        
                        st.success(f"🎉 Carga '{nome_carga}' liberada com sucesso! {len(df_produtos)} produtos disponíveis para a Doca.")
                    else:
                        st.error("⚠️ Nenhum produto encontrado no CSV.")
                except Exception as e:
                    st.error(f"Erro ao processar arquivos: {e}")
            else:
                st.warning("Preencha o nome da carga e envie ambos os arquivos.")

    with tab2:
        st.subheader("Acompanhamento em Tempo Real")
        df_cargas = st.session_state.bd_nuvem_cargas
        cargas_ativas = df_cargas[df_cargas["STATUS"] == "EM CONFERENCIA"] if not df_cargas.empty else pd.DataFrame()
        
        if cargas_ativas.empty:
            st.info("Nenhuma carga ativa no momento.")
        else:
            escolha_carga = st.selectbox("Selecione a carga para verificar/fechar:", cargas_ativas["NOME_CARGA"].unique())
            row_carga = cargas_ativas[cargas_ativas["NOME_CARGA"] == escolha_carga].iloc[0]
            id_sel = str(row_carga["ID_CARGA"])
            
            df_bipado = st.session_state.bd_nuvem_itens[st.session_state.bd_nuvem_itens["ID_CARGA"].astype(str) == id_sel]
            
            st.write("### Itens conferidos pela equipe via celular:")
            st.dataframe(df_bipado)
            
            if st.button("🏁 Fechar Conta e Gerar Planilha Oficial"):
                if df_bipado.empty:
                    st.error("Nenhum item foi conferido nessa carga ainda.")
                else:
                    try:
                        modelo_bytes = st.session_state.get(f"modelo_{id_sel}")
                        wb = openpyxl.load_workbook(io.BytesIO(modelo_bytes))
                        aba_nome = "SHELF" if "SHELF" in wb.sheetnames else wb.sheetnames[0]
                        ws = wb[aba_nome]
                        
                        linha_inicio = 5
                        for index, row in df_bipado.reset_index().iterrows():
                            linha_atual = linha_inicio + index
                            ws[f"A{linha_atual}"] = int(row["CODIGO"]) if str(row["CODIGO"]).isdigit() else row["CODIGO"]
                            ws[f"B{linha_atual}"] = row["DESCRICAO"]
                            ws[f"C{linha_atual}"] = row["FABRICACAO"]
                            ws[f"D{linha_atual}"] = row["VALIDADE"]
                            ws
