import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import re

st.set_page_config(page_title="Controle de Recebimento", layout="wide")

# CONEXÃO COM O GOOGLE SHEETS
# O Streamlit vai puxar os dados de acesso direto do arquivo 'secrets.toml' que configuraremos abaixo.
try:
    conn = st.connection("gsheets", type=st.Connection)
except Exception:
    conn = None

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
                if conn is None:
                    st.error("⚠️ O banco de dados do Google Sheets não está configurado no Streamlit Cloud. Siga o passo a passo abaixo.")
                else:
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
                            
                            # 1. Salva a carga na aba 'CARGAS' do Google Sheets para o celular ler
                            df_cargas_atual = conn.read(worksheet="CARGAS", ttl=0)
                            nova_carga = pd.DataFrame([{"ID_CARGA": id_carga, "NOME_CARGA": nome_carga, "CONTEUDO_CSV": df_produtos.to_json(orient="records"), "STATUS": "EM CONFERENCIA"}])
                            df_cargas_nova = pd.concat([df_cargas_atual, nova_carga], ignore_index=True)
                            conn.update(worksheet="CARGAS", data=df_cargas_nova)
                            
                            # Guarda o modelo Excel temporariamente na sessão do PC para o fechamento
                            st.session_state[f"modelo_{id_carga}"] = arquivo_modelo.read()
                            
                            st.success(f"🎉 Carga '{nome_carga}' gravada no Google Sheets e liberada para o celular!")
                        else:
                            st.error("⚠️ Nenhum produto encontrado no CSV.")
                    except Exception as e:
                        st.error(f"Erro ao processar arquivos: {e}")
            else:
                st.warning("Preencha o nome da carga e envie ambos os arquivos.")

    with tab2:
        st.subheader("Acompanhamento em Tempo Real")
        if conn is not None:
            try:
                df_cargas = conn.read(worksheet="CARGAS", ttl=0)
                cargas_ativas = df_cargas[df_cargas["STATUS"] == "EM CONFERENCIA"] if not df_cargas.empty else pd.DataFrame()
                
                if cargas_ativas.empty:
                    st.info("Nenhuma carga ativa no momento.")
                else:
                    escolha_carga = st.selectbox("Selecione a carga para verificar/fechar:", cargas_ativas["NOME_CARGA"].unique())
                    row_carga = cargas_ativas[cargas_ativas["NOME_CARGA"] == escolha_carga].iloc[0]
                    id_sel = str(row_carga["ID_CARGA"])
                    
                    # Puxa os bipes reais vindos do Google Sheets aba 'ITENS_CONFERIDOS'
                    df_itens_bd = conn.read(worksheet="ITENS_CONFERIDOS", ttl=0)
                    df_bipado = df_itens_bd[df_itens_bd["ID_CARGA"].astype(str) == id_sel] if not df_itens_bd.empty else pd.DataFrame()
                    
                    st.write("### Itens conferidos pela equipe via celular:")
                    st.dataframe(df_bipado)
                    
                    if st.button("🏁 Fechar Conta e Gerar Planilha Oficial"):
                        if df_bipado.empty:
                            st.error("Nenhum item foi conferido nessa carga ainda.")
                        else:
                            modelo_bytes = st.session_state.get(f"modelo_{id_sel}")
                            if not modelo_bytes:
                                st.error("Arquivo modelo não encontrado na sessão atual deste PC. Suba o modelo novamente na aba 1.")
                            else:
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
                                    ws[f"H{linha_atual}"] = row["QUANTIDADE"]
                                    
                                    ws[f"E{linha_atual}"] = f"=D{linha_atual}-TODAY()"
                                    ws[f"F{linha_atual}"] = f"=D{linha_atual}-C{linha_atual}"
                                    ws[f"G{linha_atual}"] = f"=E{linha_atual}/F{linha_atual}"
                                
                                buffer = io.BytesIO()
                                wb.save(buffer)
                                buffer.seek(0)
                                
                                # Atualiza o status da carga para FINALIZADO no Sheets
                                df_cargas.loc[df_cargas["ID_CARGA"].astype(str) == id_sel, "STATUS"] = "FINALIZADO"
                                conn.update(worksheet="CARGAS", data=df_cargas)
                                
                                st.success("🎉 Planilha gerada com sucesso!")
                                st.download_button(
                                    label="📥 Baixar Planilha de Controle Pronta",
                                    data=buffer,
                                    file_name=f"CONTROLE_SHELF_{escolha_carga.replace(' ', '_')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
            except Exception as e:
                st.error(f"Erro ao conectar com o painel: {e}")

# ----------------- 2. TELA DO CONFERENTE (CELULAR) -----------------
else:
    st.title("📱 Conferência de Entrada - Doca")
    
    if conn is None:
        st.error("Banco de dados não configurado.")
    else:
        try:
            df_cargas = conn.read(worksheet="CARGAS", ttl=0)
            cargas_disponiveis = df_cargas[df_cargas["STATUS"] == "EM CONFERENCIA"] if not df_cargas.empty else pd.DataFrame()
            
            if cargas_disponiveis.empty:
                st.success("✅ Nenhuma carga pendente de conferência na doca!")
                if st.button("🔄 Atualizar Página"):
                    st.rerun()
            else:
                carga_selecionada = st.selectbox("Selecione a Carga para conferir:", cargas_disponiveis["NOME_CARGA"].unique())
                row_c = cargas_disponiveis[cargas_disponiveis["NOME_CARGA"] == carga_selecionada].iloc[0]
                id_carga_ativa = str(row_c["ID_CARGA"])
                
                df_produtos_carga = pd.read_json(io.StringIO(row_c["CONTEUDO_CSV"]))
                
                st.markdown("---")
                conferente = st.text_input("Nome do Conferente:")
                codigo_bipado = st.text_input("Digite ou Bipe o Código do Produto:").strip().lstrip('0')
                
                if codigo_bipado:
                    df_produtos_carga["Codigo_Prod"] = df_produtos_carga["Codigo_Prod"].astype(str)
                    item = df_produtos_carga[df_produtos_carga["Codigo_Prod"] == codigo_bipado]
                    
                    if not item.empty:
                        desc_item = item.iloc[0]["Descricao_Prod"]
                        st.info(f"📦 **Item:** {desc_item}")
                        
                        with st.form(key="form_celular", clear_on_submit=True):
                            col1, col2 = st.columns(2)
                            with col1:
                                f_fab = st.date_input("Data de Fabricação", value=datetime.today())
                            with col2:
                                f_ven = st.date_input("Data de Vencimento", value=datetime.today())
                            
                            f_qtd = st.number_input("Quantidade:", min_value=1, step=1)
                            btn_enviar = st.form_submit_button("🚀 Enviar Lote para o Servidor")
                            
                            if btn_enviar:
                                if conferente:
                                    # Grava o lote inserido direto na aba 'ITENS_CONFERIDOS' do Sheets
                                    df_itens_atual = conn.read(worksheet="ITENS_CONFERIDOS", ttl=0)
                                    novo_lote = pd.DataFrame([{
                                        "ID_CARGA": id_carga_ativa,
                                        "CODIGO": codigo_bipado,
                                        "DESCRICAO": desc_item,
                                        "FABRICACAO": f_fab.strftime('%d/%m/%Y'),
                                        "VALIDADE": f_ven.strftime('%d/%m/%Y'),
                                        "QUANTIDADE": f_qtd,
                                        "CONFERENTE": conferente
                                    }])
                                    df_itens_novo = pd.concat([df_itens_atual, novo_lote], ignore_index=True)
                                    conn.update(worksheet="ITENS_CONFERIDOS", data=df_itens_novo)
                                    st.success("✔️ Lote enviado com sucesso para a nuvem!")
                                else:
                                    st.error("⚠️ Preencha seu nome antes de enviar.")
                    else:
                        st.error("❌ Código de produto não encontrado nesta carga.")
        except Exception as e:
            st.error(f"Erro ao carregar dados na doca: {e}")
