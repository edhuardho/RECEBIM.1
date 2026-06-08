import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import re

st.set_page_config(page_title="Controle de Recebimento", layout="wide")

# COLOQUE O LINK DA SUA PLANILHA DO GOOGLE SHEETS AQUI DENTRO DAS ASPAS (OPCIONAL PARA ESTA VERSÃO):
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit?usp=sharing"

# Inicialização dos bancos de dados temporários no servidor do Streamlit
if "bd_simulado_cargas" not in st.session_state:
    st.session_state.bd_simulado_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
if "bd_simulado_itens" not in st.session_state:
    st.session_state.bd_simulado_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])

# ----------------- INTERFACE SHIFT (MENU) -----------------
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
                    # Lendo o arquivo primeiro como texto para evitar o erro de 'errors' no read_csv
                    conteudo = arquivo_csv.read().decode("utf-8", errors="ignore")
                    
                    # Converte o texto puro em um DataFrame legível separando por linhas
                    linhas_texto = conteudo.splitlines()
                    dados_linhas = [linha.split(',') for linha in linhas_texto if linha.strip()]
                    
                    df_bruto = pd.DataFrame(dados_linhas)
                    
                    linhas_produtos = []
                    
                    for idx, row in df_bruto.iterrows():
                        if len(row) == 0:
                            continue
                            
                        # Pega o conteúdo da primeira célula da linha
                        texto_linha = str(row.iloc[0]).strip()
                        
                        # Se o arquivo tiver mais colunas (separado por vírgula), junta a descrição
                        if len(row) > 1 and row.iloc[1] is not None and str(row.iloc[1]).strip() != "":
                            desc_extra = str(row.iloc[1]).strip()
                            # Se a primeira coluna for só número, ela é o código e a segunda é a descrição
                            if texto_linha.isdigit():
                                codigo = texto_linha.lstrip('0')
                                descricao = desc_extra
                                if not any(t in descricao.upper() for t in ["TOTAL", "EMISSÃO", "PÁGINA", "RELATÓRIO"]):
                                    linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})
                                continue

                        # Caso o código e a descrição estejam na mesma célula separados por " - "
                        if " - " in texto_linha:
                            texto_limpo = texto_linha.replace('"', '').strip()
                            partes = texto_limpo.split(' - ', 1)
                            if len(partes) == 2:
                                codigo = partes[0].strip().lstrip('0')
                                descricao = partes[1].strip()
                                if codigo.isdigit() and not any(t in descricao.upper() for t in ["TOTAL", "EMISSÃO", "PÁGINA", "RELATÓRIO"]):
                                    linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})

# ----------------- 2. TELA DO CONFERENTE (CELULAR) -----------------
else:
    st.title("📱 Conferência de Entrada - Doca")
    st.write("Selecione a carga liberada pelo supervisor, digite o código e insira os dados do lote.")
    
    if st.session_state.bd_simulado_cargas.empty:
        st.warning("⚠️ Nenhuma carga foi liberada pelo supervisor no momento. Aguarde a liberação.")
    else:
        cargas_disponiveis = st.session_state.bd_simulado_cargas[st.session_state.bd_simulado_cargas["STATUS"] == "EM CONFERENCIA"]
        
        if cargas_disponiveis.empty:
            st.success("✅ Todas as cargas do dia já foram conferidas!")
        else:
            carga_selecionada = st.selectbox("Selecione a Carga que chegou:", cargas_disponiveis["NOME_CARGA"].unique())
            row_c = cargas_disponiveis[cargas_disponiveis["NOME_CARGA"] == carga_selecionada].iloc[0]
            id_carga_ativa = row_c["ID_CARGA"]
            
            # Recria o DataFrame lendo o formato string JSON correto
            df_produtos_carga = pd.read_json(io.StringIO(row_c["CONTEUDO_CSV"]))
            
            st.markdown("---")
            conferente = st.text_input("Nome do Conferente:", key="nome_conf")
            codigo_bipado = st.text_input("Digite ou Bipe o Código do Produto:", key="code_bip").strip().lstrip('0')
            
            if codigo_bipado:
                # Converte coluna para string para bater com o input text
                df_produtos_carga["Codigo_Prod"] = df_produtos_carga["Codigo_Prod"].astype(str)
                item = df_produtos_carga[df_produtos_carga["Codigo_Prod"] == codigo_bipado]
                
                if not item.empty:
                    desc_item = item.iloc[0]["Descricao_Prod"]
                    st.info(f"📦 **Produto:** {desc_item}")
                    
                    with st.form(key="form_celular", clear_on_submit=True):
                        col1, col2 = st.columns(2)
                        with col1:
                            f_fab = st.date_input("Data de Fabricação", value=datetime.today())
                        with col2:
                            f_ven = st.date_input("Data de Vencimento", value=datetime.today())
                        
                        f_qtd = st.number_input("Quantidade (Caixas/Unidades):", min_value=1, step=1)
                        btn_enviar = st.form_submit_button("🚀 Enviar Lote para o PC")
                        
                        if btn_enviar:
                            if conferente:
                                novo_lote = pd.DataFrame([{
                                    "ID_CARGA": id_carga_ativa,
                                    "CODIGO": codigo_bipado,
                                    "DESCRICAO": desc_item,
                                    "FABRICACAO": f_fab.strftime('%d/%m/%Y'),
                                    "VALIDADE": f_ven.strftime('%d/%m/%Y'),
                                    "QUANTIDADE": f_qtd,
                                    "CONFERENTE": conferente
                                }])
                                
                                st.session_state.bd_simulado_itens = pd.concat([st.session_state.bd_simulado_itens, novo_lote], ignore_index=True)
                                st.success("✔️ Lote enviado com sucesso para o painel do supervisor!")
                            else:
                                st.error("⚠️ Digite seu nome antes de salvar o lote.")
                else:
                    st.error("❌ Código de produto não pertence a esta carga.")
            
            st.markdown("---")
            st.subheader("📋 Meus Itens Enviados nesta Carga")
            meus_itens = st.session_state.bd_simulado_itens[st.session_state.bd_simulado_itens["ID_CARGA"] == id_carga_ativa]
            if not meus_itens.empty:
                st.dataframe(meus_itens[["CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE"]])
