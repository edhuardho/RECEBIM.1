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
                    # Lendo o arquivo primeiro como texto puro para limpar quebras e recuos complexos
                    conteudo = arquivo_csv.read().decode("utf-8", errors="ignore")
                    linhas_texto = conteudo.splitlines()
                    
                    linhas_produtos = []
                    
                    for linha in linhas_texto:
                        if not linha.strip():
                            continue
                        
                        # Limpa caracteres de estrutura e aspas que poluem o relatório
                        linha_limpa = linha.replace('"', '').replace('-[-[-[-[-[-[-[-[', '').replace('\t', ' ').strip()
                        
                        # Captura códigos numéricos de 4 a 8 dígitos que venham seguidos de " - " e a descrição
                        match = re.search(r'(\d{4,8})\s*-\s*([^,]+)', linha_limpa)
                        
                        if match:
                            codigo = match.group(1).strip().lstrip('0')
                            descricao = match.group(2).strip()
                            
                            # Filtro de segurança contra cabeçalhos ou linhas de controle
                            if not any(t in descricao.upper() for t in ["TOTAL", "EMISSÃO", "PÁGINA", "RELATÓRIO", "CLIENTE", "MOTORISTA"]):
                                linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})
                    
                    if len(linhas_produtos) > 0:
                        df_produtos = pd.DataFrame(linhas_produtos).drop_duplicates(subset=['Codigo_Prod'])
                        
                        # Gera ID única baseada no timestamp atual
                        id_carga = str(int(datetime.now().timestamp()))
                        nova_linha = pd.DataFrame([{"ID_CARGA": id_carga, "NOME_CARGA": nome_carga, "CONTEUDO_CSV": df_produtos.to_json(orient="records"), "STATUS": "EM CONFERENCIA"}])
                        st.session_state.bd_simulado_cargas = pd.concat([st.session_state.bd_simulado_cargas, nova_linha], ignore_index=True)
                        
                        # Salva o arquivo de modelo padrão associado à carga
                        st.session_state[f"modelo_{id_carga}"] = arquivo_modelo.read()
                        
                        st.success(f"🎉 Carga '{nome_carga}' liberada com sucesso! {len(df_produtos)} produtos disponíveis para a Doca.")
                    else:
                        st.error("⚠️ Estrutura de dados não reconhecida. Certifique-se de que o CSV é o relatório correto.")
                
                except Exception as e:
                    st.error(f"Erro ao processar arquivos: {e}")
            else:
                st.warning("Preencha o nome da carga e envie ambos os arquivos para liberar.")

    with tab2:
        st.subheader("Acompanhamento em Tempo Real")
        if st.session_state.bd_simulado_cargas.empty:
            st.info("Nenhuma carga ativa ou em andamento no momento.")
        else:
            cargas_ativas = st.session_state.bd_simulado_cargas[st.session_state.bd_simulado_cargas["STATUS"] == "EM CONFERENCIA"]
            
            if cargas_ativas.empty:
                st.info("Todas as cargas liberadas já foram finalizadas.")
            else:
                escolha_carga = st.selectbox("Selecione a carga para verificar/fechar:", cargas_ativas["NOME_CARGA"].unique())
                row_carga = cargas_ativas[cargas_ativas["NOME_CARGA"] == escolha_carga].iloc[0]
                id_sel = row_carga["ID_CARGA"]
                
                # Puxa os itens bipados do celular para o painel do supervisor em tempo real
                df_bipado = st.session_state.bd_simulado_itens[st.session_state.bd_simulado_itens["ID_CARGA"] == id_sel]
                
                st.write("### Itens já conferidos pela equipe na doca:")
                st.dataframe(df_bipado)
                
                if st.button("🏁 Fechar Conta e Gerar Planilha Oficial"):
                    if df_bipado.empty:
                        st.error("Nenhum item foi conferido nessa carga ainda pelo celular.")
                    else:
                        try:
                            # Monta o arquivo oficial de saída injetando as fórmulas nas células corretas
                            modelo_bytes = st.session_state.get(f"modelo_{id_sel}")
                            wb = openpyxl.load_workbook(io.BytesIO(modelo_bytes))
                            aba_
