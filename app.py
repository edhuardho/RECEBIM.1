import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import re

st.set_page_config(page_title="Controle de Recebimento", layout="wide")

# Inicialização dos bancos de dados temporários no servidor do Streamlit
if "bd_simulado_cargas" not in st.session_state:
    st.session_state.bd_simulado_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
if "bd_simulado_itens" not in st.session_state:
    st.session_state.bd_simulado_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])

# ----------------- INTERFACE SHIFT (MENU LATERAL) -----------------
st.sidebar.title("🎮 Controle de Acesso")
perfil = st.sidebar.radio("Selecione o seu Perfil:", ["🖥️ Painel do Supervisor (PC)", "📱 Conferência na Doca (Celular)"])

# Botão master para resetar o servidor caso queira limpar o dia de trabalho
st.sidebar.markdown("---")
if st.sidebar.button("🧹 Resetar Todas as Cargas do Sistema"):
    st.session_state.bd_simulado_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
    st.session_state.bd_simulado_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])
    st.rerun()

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
                    # Lendo o arquivo primeiro como texto para evitar problemas de encoding e colunas
                    conteudo = arquivo_csv.read().decode("utf-8", errors="ignore")
                    linhas_texto = conteudo.splitlines()
                    
                    # Cria a matriz separando por vírgulas (padrão do arquivo Tirolez)
                    dados_linhas = []
                    for linha in linhas_texto:
                        if linha.strip():
                            # Limpa aspas duplas que o CSV costuma trazer envolvidas no texto
                            linha_limpa = linha.replace('"', '')
                            dados_linhas.append(linha_limpa.split(','))
                    
                    df_bruto = pd.DataFrame(dados_linhas)
                    linhas_produtos = []
                    
                    for idx, row in df_bruto.iterrows():
                        if row.dropna().empty:
                            continue
                            
                        # Pega o conteúdo da primeira célula da linha
                        texto_linha = str(row.iloc[0]).strip()
