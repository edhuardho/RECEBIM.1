import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io
import re

st.set_page_config(page_title="Controle de Recebimento", layout="wide")

# 🖥️ BANCO DE DADOS GLOBAL COMPARTILHADO (A nível de Servidor)
# Isso garante que o celular enxergue exatamente a mesma carga que o PC criou!
if "bd_global_cargas" not in st.session_state.__class__.__dict__:
    st.session_state.__class__.bd_global_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
if "bd_global_itens" not in st.session_state.__class__.__dict__:
    st.session_state.__class__.bd_global_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])
if "modelos_excel_memoria" not in st.session_state.__class__.__dict__:
    st.session_state.__class__.modelos_excel_memoria = {}

# ----------------- INTERFACE SHIFT (MENU LATERAL) -----------------
st.sidebar.title("🎮 Controle de Acesso")
perfil = st.sidebar.radio("Selecione o seu Perfil:", ["🖥️ Painel do Supervisor (PC)", "📱 Conferência na Doca (Celular)"])

# Botão master para resetar o servidor caso queira limpar o dia de trabalho
st.sidebar.markdown("---")
if st.sidebar.button("🧹 Limpar Todas as Cargas do Sistema"):
    st.session_state.__class__.bd_global_cargas = pd.DataFrame(columns=["ID_CARGA", "NOME_CARGA", "CONTEUDO_CSV", "STATUS"])
    st.session_state.__class__.bd_global_itens = pd.DataFrame(columns=["ID_CARGA", "CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE", "CONFERENTE"])
    st.session_state.__class__.modelos_excel_memoria = {}
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
                    conteudo = arquivo_csv.read().decode("utf-8", errors="ignore")
                    linhas_texto = conteudo.splitlines()
                    
                    linhas_produtos = []
                    for linha in linhas_texto:
                        if not linha.strip():
                            continue
                        linha_limpa = line_clean = linha.replace('"', '').replace('-[-[-[-[-[-[-[-[', '').replace('\t', ' ').strip()
                        match = re.search(r'(\d{4,8})\s*-\s*([^,]+)', linha_limpa)
                        
                        if match:
                            codigo = match.group(1).strip().lstrip('0')
                            descricao = match.group(2).strip()
                            if not any(t in descricao.upper() for t in ["TOTAL", "EMISSÃO", "PÁGINA", "RELATÓRIO", "CLIENTE", "MOTORISTA"]):
                                linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})
                    
                    if len(linhas_produtos) > 0:
                        df_produtos = pd.DataFrame(linhas_produtos).drop_duplicates(subset=['Codigo_Prod'])
                        id_carga = str(int(datetime.now().timestamp()))
                        
                        # Injeta no banco global compartilhado do servidor
                        nova_carga = pd.DataFrame([{"ID_CARGA": id_carga, "NOME_CARGA": nome_carga, "CONTEUDO_CSV": df_produtos.to_json(orient="records"), "STATUS": "EM CONFERENCIA"}])
                        st.session_state.__class__.bd_global_cargas = pd.concat([st.session_state.__class__.bd_global_cargas, nova_carga], ignore_index=True)
                        
                        # Guarda o binário do modelo Excel associado a ID
                        st.session_state.__class__.modelos_excel_memoria[id_carga] = arquivo_modelo.read()
                        
                        st.success(f"🎉 Carga '{nome_carga}' liberada com sucesso! {len(df_produtos)} produtos disponíveis para a Doca no Celular.")
                    else:
                        st.error("⚠️ Nenhum produto encontrado no CSV.")
                except Exception as e:
                    st.error(f"Erro ao processar arquivos: {e}")
            else:
                st.warning("Preencha o nome da carga e envie ambos os arquivos.")

    with tab2:
        st.subheader("Acompanhamento em Tempo Real")
        df_cargas = st.session_state.__class__.bd_global_cargas
        cargas_ativas = df_cargas[df_cargas["STATUS"] == "EM CONFERENCIA"] if not df_cargas.empty else pd.DataFrame()
        
        if cargas_ativas.empty:
            st.info("Nenhuma carga ativa no momento.")
        else:
            escolha_carga = st.selectbox("Selecione a carga para verificar/fechar:", cargas_ativas["NOME_CARGA"].unique())
            row_carga = cargas_ativas[cargas_ativas["NOME_CARGA"] == escolha_carga].iloc[0]
            id_sel = str(row_carga["ID_CARGA"])
            
            df_bipado = st.session_state.__class__.bd_global_itens[st.session_state.__class__.bd_global_itens["ID_CARGA"].astype(str) == id_sel]
            
            st.write("### Itens conferidos pela equipe via celular:")
            st.dataframe(df_bipado)
            
            if st.button("🏁 Fechar Conta e Gerar Planilha Oficial"):
                if df_bipado.empty:
                    st.error("Nenhum item foi conferido nessa carga ainda.")
                else:
                    try:
                        modelo_bytes = st.session_state.__class__.modelos_excel_memoria.get(id_sel)
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
                        
                        # Altera o status para Finalizado
                        st.session_state.__class__.bd_global_cargas.loc[st.session_state.__class__.bd_global_cargas["ID_CARGA"].astype(str) == id_sel, "STATUS"] = "FINALIZADO"
                        
                        st.success("🎉 Planilha gerada com sucesso!")
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
    
    # Menu rápido de atualização no topo da tela do celular
    if st.button("🔄 Atualizar Lista de Cargas"):
        st.rerun()
        
    df_cargas = st.session_state.__class__.bd_global_cargas
    cargas_disponiveis = df_cargas[df_cargas["STATUS"] == "EM CONFERENCIA"] if not df_cargas.empty else pd.DataFrame()
    
    if cargas_disponiveis.empty:
        st.success("✅ Nenhuma carga pendente de conferência na doca!")
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
                    btn_enviar = st.form_submit_button("🚀 Enviar Lote para o Supervisor")
                    
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
                            st.session_state.__class__.bd_global_itens = pd.concat([st.session_state.__class__.bd_global_itens, novo_lote], ignore_index=True)
                            st.success("✔️ Lote enviado com sucesso para o painel do supervisor!")
                        else:
                            st.error("⚠️ Preencha seu nome antes de enviar.")
            else:
                st.error("❌ Código de produto não encontrado nesta carga.")
        
        st.markdown("---")
        st.subheader("📋 Meus Itens Enviados")
        meus_itens = st.session_state.__class__.bd_global_itens[st.session_state.__class__.bd_global_itens["ID_CARGA"] == id_carga_ativa]
        if not meus_itens.empty:
            st.dataframe(meus_itens[["CODIGO", "DESCRICAO", "FABRICACAO", "VALIDADE", "QUANTIDADE"]])
