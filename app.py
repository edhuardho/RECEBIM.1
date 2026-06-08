import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
import io

# Configuração da página para celular e computador
st.set_page_config(page_title="Controle de Recebimento", layout="centered")

st.title("📦 Sistema de Recebimento & Shelf Life")
st.write("Insira o CSV da Delicari, faça a conferência e gere a planilha oficial.")

# Inicialização das variáveis de sessão
if "dados_conferidos" not in st.session_state:
    st.session_state.dados_conferidos = []
if "csv_tratado" not in st.session_state:
    st.session_state.csv_tratado = None

# --- 1. UPLOAD DOS ARQUIVOS ---
st.subheader("📁 1. Upload de Arquivos")
arquivo_csv = st.file_uploader("Suba o arquivo CSV bruto (DELICARI)", type=["csv"], key="csv")
arquivo_modelo = st.file_uploader("Suba o seu modelo padrão (SHELF - PADRÃO.xlsx)", type=["xlsx"], key="modelo")

if arquivo_csv is not None and st.session_state.csv_tratado is None:
    try:
        # Lê o CSV da Delicari separando por ponto e vírgula
        df_bruto = pd.read_csv(arquivo_csv, sep=";", encoding="utf-8")
        if 'Produto' in df_bruto.columns:
            # Separa Código e Descrição
            df_bruto[['Codigo_Prod', 'Descricao_Prod']] = df_bruto['Produto'].astype(str).str.split('-', n=1, expand=True)
            df_bruto['Codigo_Prod'] = df_bruto['Codigo_Prod'].str.strip()
            df_bruto['Descricao_Prod'] = df_bruto['Descricao_Prod'].str.strip()
            st.session_state.csv_tratado = df_bruto
            st.success("✅ CSV processado com sucesso!")
        else:
            st.error("Coluna 'Produto' não encontrada no CSV.")
    except Exception as e:
        st.error(f"Erro ao processar CSV: {e}")

# --- 2. TELA DE CONFERÊNCIA (CELULAR / PC) ---
if st.session_state.csv_tratado is not None and arquivo_modelo is not None:
    st.markdown("---")
    st.subheader("📝 2. Conferência de Itens na Doca")
    
    df_atual = st.session_state.csv_tratado
    codigo_digitado = st.text_input("Digite ou Bipe o Código do Produto:").strip()

    if codigo_digitado:
        item_encontrado = df_atual[df_atual['Codigo_Prod'] == codigo_digitado]

        if not item_encontrado.empty:
            descricao = item_encontrado.iloc[0]['Descricao_Prod']
            qtd_esperada = item_encontrado['Qtde Atendida'].sum()

            st.info(f"📦 **Produto:** {descricao} \n\n 📊 **Qtd Esperada no CSV:** {qtd_esperada}")

            # Formulário para múltiplos vencimentos
            with st.form(key="form_lotes", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    data_fab = st.date_input("Data de Fabricação", value=datetime.today())
                with col2:
                    data_ven = st.date_input("Data de Vencimento", value=datetime.today())
                
                qtd_lote = st.number_input("Quantidade deste Lote:", min_value=1, step=1)
                btn_salvar = st.form_submit_button("➕ Salvar Lote")

                if btn_salvar:
                    st.session_state.dados_conferidos.append({
                        "CODIGO": codigo_digitado,
                        "DESCRIÇÃO DO PRODUTO": descricao,
                        "FABRICAÇÃO": data_fab.strftime('%d/%m/%Y'),
                        "VALIDADE": data_ven.strftime('%d/%m/%Y'),
                        "QUANTIDADE": qtd_lote
                    })
                    st.toast(f"Lote salvo para o item {codigo_digitado}!", icon="✔️")
        else:
            st.error("⚠️ Código não consta neste CSV de recebimento.")

    # --- 3. EXIBIÇÃO E EXPORTAÇÃO ---
    if st.session_state.dados_conferidos:
        st.markdown("---")
        st.subheader("📋 Itens Conferidos")
        df_conferidos = pd.DataFrame(st.session_state.dados_conferidos)
        st.dataframe(df_conferidos)

        # Botão para processar e gerar o download
        if st.button("🏁 Finalizar Recebimento e Criar Planilha"):
            try:
                # Carrega o modelo original usando openpyxl
                wb = openpyxl.load_workbook(arquivo_modelo)
                
                # Seleciona a aba SHELF (garante compatibilidade de letras maiúsculas)
                aba_nome = "SHELF" if "SHELF" in wb.sheetnames else wb.sheetnames[0]
                ws = wb[aba_nome]
                
                # Começa a preencher a partir da linha 2 (abaixo do cabeçalho)
                linha_inicio = 2
                
                for index, row in df_conferidos.iterrows():
                    linha_atual = linha_inicio + index
                    
                    # Escreve os dados coletados nas colunas corretas
                    ws[f"A{linha_atual}"] = int(row["CODIGO"]) if row["CODIGO"].isdigit() else row["CODIGO"]
                    ws[f"B{linha_atual}"] = row["DESCRIÇÃO DO PRODUTO"]
                    ws[f"C{linha_atual}"] = row["FABRICAÇÃO"]
                    ws[f"D{linha_atual}"] = row["VALIDADE"]
                    ws[f"H{linha_atual}"] = row["QUANTIDADE"]
                    
                    # Injeta as fórmulas originais do seu Excel adaptadas para a linha atual
                    ws[f"E{linha_atual}"] = f"=D{linha_atual}-TODAY()"
                    ws[f"F{linha_atual}"] = f"=D{linha_atual}-C{linha_atual}"
                    ws[f"G{linha_atual}"] = f"=E{linha_atual}/F{linha_atual}"
                
                # Salva o arquivo em memória para disponibilizar o download
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.success("🎉 Planilha oficial gerada com sucesso!")
                st.download_button(
                    label="📥 Baixar Planilha de Controle Pronta",
                    data=buffer,
                    file_name=f"CONTROLE_SHELF_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao gerar o arquivo final: {e}")