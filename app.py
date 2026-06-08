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
                    # Lendo o CSV tratando a separação por vírgula que vem no arquivo da Tirolez
                    # 'header=None' garante que não vamos perder a primeira linha de produto
                    df_bruto = pd.read_csv(arquivo_csv, sep=None, engine='python', header=None, encoding='utf-8', errors='ignore')
                    
                    linhas_produtos = []
                    
                    for idx, row in df_bruto.iterrows():
                        # Pega o conteúdo da primeira célula da linha
                        texto_linha = str(row.iloc[0]).strip()
                        
                        # Se o arquivo tiver mais colunas (separado por vírgula), junta a descrição
                        if len(row) > 1 and not pd.isna(row.iloc[1]):
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
                    
                    if len(linhas_produtos) > 0:
                        df_produtos = pd.DataFrame(linhas_produtos).drop_duplicates(subset=['Codigo_Prod'])
                        
                        id_carga = str(int(datetime.now().timestamp()))
                        nova_linha = pd.DataFrame([{"ID_CARGA": id_carga, "NOME_CARGA": nome_carga, "CONTEUDO_CSV": df_produtos.to_json(orient="records"), "STATUS": "EM CONFERENCIA"}])
                        st.session_state.bd_simulado_cargas = pd.concat([st.session_state.bd_simulado_cargas, nova_linha], ignore_index=True)
                        
                        # Guarda o modelo Excel na memória do servidor
                        st.session_state[f"modelo_{id_carga}"] = arquivo_modelo.read()
                        
                        st.success(f"🎉 Carga '{nome_carga}' liberada com sucesso! {len(df_produtos)} produtos disponíveis para a Doca.")
                    else:
                        st.error("⚠️ Estrutura de dados não reconhecida. Verifique se o arquivo enviado é o correto.")
                except Exception as e:
                    st.error(f"Erro ao processar arquivos: {e}")
            else:
                st.warning("Preencha o nome da carga e envie ambos os arquivos.")

    with tab2:
        st.subheader("Acompanhamento em Tempo Real")
        if st.session_state.bd_simulado_cargas.empty:
            st.info("Nenhuma carga ativa ou em andamento no momento.")
        else:
            cargas_ativas = st.session_state.bd_simulado_cargas[st.session_state.bd_simulado_cargas["STATUS"] == "EM CONFERENCIA"]
            
            if cargas_ativas.empty:
                st.info("Todas as cargas foram finalizadas.")
            else:
                escolha_carga = st.selectbox("Selecione a carga para verificar/fechar:", cargas_ativas["NOME_CARGA"].unique())
                row_carga = cargas_ativas[cargas_ativas["NOME_CARGA"] == escolha_carga].iloc[0]
                id_sel = row_carga["ID_CARGA"]
                
                df_bipado = st.session_state.bd_simulado_itens[st.session_state.bd_simulado_itens["ID_CARGA"] == id_sel]
                
                st.write("### Itens já conferidos pela equipe na doca:")
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
                                ws[f"H{linha_atual}"] = row["QUANTIDADE"]
                                
                                ws[f"E{linha_atual}"] = f"=D{linha_atual}-TODAY()"
                                ws[f"F{linha_atual}"] = f"=D{linha_atual}-C{linha_atual}"
                                ws[f"G{linha_atual}"] = f"=E{linha_atual}/F{linha_atual}"
                            
                            buffer = io.BytesIO()
                            wb.save(buffer)
                            buffer.seek(0)
                            
                            st.session_state.bd_simulado_cargas.loc[st.session_state.bd_simulado_cargas["ID_CARGA"] == id_sel, "STATUS"] = "FINALIZADO"
                            
                            st.success("🎉 Relatório processado com sucesso!")
                            st.download_button(
                                label="📥 Baixar Planilha de Controle Pronta",
                                data=buffer,
                                file_name=f"CONTROLE_SHELF_{escolha_carga.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"Erro ao gerar arquivo: {e}")

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
